from __future__ import annotations

import asyncio
import base64
import logging
import sys
import threading
from decimal import Decimal
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from horoshop_prices import (
    ArticleMappings, CatalogIndex, Credentials, FieldChange, HoroshopClient, HoroshopPricesError, PricePlan, PriceRow, Settings, WholesaleChange,
    build_excel_template, build_failed_excel, import_results, load_article_mappings, load_settings, normalize, parse_article_mappings,
    parse_decimal, parse_excel_prices, parse_price_type, parse_threshold, plan_prices, save_article_mappings,
)


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MAPPINGS_FILE = PROJECT_DIR / "data" / "article_mappings.json"
settings: Settings | None = None
logger = logging.getLogger(__name__)
mapping_lock = threading.Lock()


class PublicLogStream:
    def __init__(self, primary: Path, fallback: Path) -> None:
        self.primary, self.fallback, self.encoding = primary, fallback, "utf-8"

    def write(self, message: str) -> int:
        for path in (self.primary, self.fallback):
            try:
                path.parent.mkdir(parents=True, exist_ok=True)
                with path.open("a", encoding=self.encoding) as file:
                    file.write(message)
                break
            except OSError:
                continue
        return len(message)

    def flush(self) -> None: pass
    def isatty(self) -> bool: return False


def get_settings() -> Settings:
    global settings
    if settings is None:
        if not CONFIG_FILE.exists():
            raise RuntimeError(f"Configuration file was not found: {CONFIG_FILE}. Create config.json from config.example.json.")
        settings = load_settings(CONFIG_FILE)
    return settings


def configure_service_output(runtime_settings: Settings) -> None:
    selected = runtime_settings.public_log_file
    fallback = PROJECT_DIR / "logs" / "horoshop_prices.log"
    try:
        selected.parent.mkdir(parents=True, exist_ok=True)
        selected.touch(exist_ok=True)
    except OSError:
        selected = fallback
        selected.parent.mkdir(parents=True, exist_ok=True)
        selected.touch(exist_ok=True)
    stream = PublicLogStream(selected, fallback)
    sys.stdout = stream
    sys.stderr = stream
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s", stream=stream, force=True)
    logger.info("Price service output is writing to %s", selected)


def credentials_from(data: dict[str, Any]) -> Credentials:
    credentials = Credentials(normalize(data.get("login")), normalize(data.get("password")), normalize(data.get("token")))
    if not credentials.token and (not credentials.login or not credentials.password):
        raise HoroshopPricesError("Вкажіть логін і пароль API або чинний токен.")
    return credentials


def catalog_for(credentials: Credentials) -> tuple[CatalogIndex, HoroshopClient]:
    client = HoroshopClient(get_settings(), credentials)
    return CatalogIndex.from_raw(client.export_catalog()), client


def serialise(plan: PricePlan, status: str | None = None, message: str = "") -> dict[str, Any]:
    row = plan.rows[0]
    changes: list[str] = []
    if row.current.value is not None:
        changes.append(f"Поточна ціна: {row.current.value}")
    if row.current_percent_of_old is not None:
        changes.append(f"Поточна ціна: {row.current_percent_of_old}% від РРЦ")
    if row.old.value is not None:
        changes.append(f"РРЦ: {row.old.value}")
    elif row.old.delete:
        changes.append("РРЦ: видалити")
    if row.move_current_to_old:
        changes.append("Поточну ціну перенести в РРЦ")
    if row.move_old_to_current:
        changes.append("РРЦ перенести в поточну ціну")
    for wholesale in row.wholesale:
        label = f"Опт {wholesale.tier}" + (f" від {wholesale.threshold} шт." if wholesale.threshold else "")
        changes.append(label + (": видалити" if wholesale.change.delete else f": {wholesale.change.value}"))
    return {
        "article": plan.display_article,
        "internal_article": plan.article,
        "row_number": row.row_number,
        "changes": changes,
        "status": status or ("ready" if plan.ready else "invalid"),
        "message": message or plan.error or "; ".join(plan.warnings),
    }


def with_failed_excel(result: dict[str, Any], failures: list[tuple[PricePlan, str, str]]) -> dict[str, Any]:
    if failures:
        result["failed_excel"] = base64.b64encode(build_failed_excel(failures)).decode("ascii")
        result["failed_excel_name"] = "horoshop_prices_failed.xlsx"
    return result


def preview_rows(rows: list[PriceRow], credentials: Credentials, normalize_articles: bool, mappings: ArticleMappings) -> dict[str, Any]:
    catalog, _ = catalog_for(credentials)
    plans = plan_prices(rows, catalog, get_settings(), normalize_articles, mappings)
    items = [serialise(plan) for plan in plans]
    failures = [(plan, "Помилка", item["message"]) for plan, item in zip(plans, items) if not plan.ready]
    return with_failed_excel({"items": items, "ready": sum(plan.ready for plan in plans), "errors": len(failures)}, failures)


def execute_rows(rows: list[PriceRow], credentials: Credentials, normalize_articles: bool = True, mappings: ArticleMappings | None = None) -> dict[str, Any]:
    catalog, client = catalog_for(credentials)
    plans = plan_prices(rows, catalog, get_settings(), normalize_articles, mappings or ArticleMappings.empty())
    payload = [plan.payload for plan in plans if plan.ready]
    api_results: dict[str, tuple[bool, str]] = {}
    for start in range(0, len(payload), get_settings().batch_size):
        api_results.update(import_results(client.import_products(payload[start:start + get_settings().batch_size])))
    items: list[dict[str, Any]] = []
    failures: list[tuple[PricePlan, str, str]] = []
    for plan in plans:
        if not plan.ready:
            item = serialise(plan, "invalid")
            items.append(item)
            failures.append((plan, "Помилка перевірки", item["message"]))
            continue
        success, message = api_results.get(plan.article, (False, "API не повернуло результат для товару."))
        joined = "; ".join(item for item in ("; ".join(plan.warnings), message) if item)
        item = serialise(plan, "synced" if success else "error", joined)
        items.append(item)
        if not success:
            failures.append((plan, "Помилка API", item["message"]))
    return with_failed_excel({"items": items, "imported": sum(item["status"] == "synced" for item in items), "errors": sum(item["status"] in {"error", "invalid"} for item in items)}, failures)


def parse_bool(value: Any, default: bool = False) -> bool:
    text = normalize(value).casefold()
    if not text:
        return default
    return text not in {"0", "false", "ні", "нет", "no", "off"}


def stored_mappings() -> ArticleMappings:
    with mapping_lock:
        return load_article_mappings(MAPPINGS_FILE)


def add_to_stored_mappings(uploaded_mappings: ArticleMappings) -> ArticleMappings:
    with mapping_lock:
        combined = load_article_mappings(MAPPINGS_FILE).merged_with(uploaded_mappings)
        save_article_mappings(MAPPINGS_FILE, combined)
        return combined


async def uploaded_mappings(form: Any) -> ArticleMappings | None:
    mapping_file = form.get("mapping_file")
    if mapping_file is None or not hasattr(mapping_file, "read") or not normalize(getattr(mapping_file, "filename", "")):
        return None
    if not str(getattr(mapping_file, "filename", "")).lower().endswith((".xlsx", ".xlsm")):
        raise HoroshopPricesError("Файл сопоставлень має бути .xlsx або .xlsm.")
    mapping_contents = await mapping_file.read()
    if not mapping_contents or len(mapping_contents) > MAX_UPLOAD_BYTES:
        raise HoroshopPricesError("Файл сопоставлень порожній або перевищує 20 МБ.")
    return await asyncio.to_thread(parse_article_mappings, mapping_contents)


async def uploaded_rows(request: Request) -> tuple[list[PriceRow], Credentials, bool, ArticleMappings]:
    form = await request.form()
    uploaded = form.get("file")
    if uploaded is None or not hasattr(uploaded, "read"):
        raise HoroshopPricesError("Оберіть Excel-файл .xlsx або .xlsm.")
    if not str(getattr(uploaded, "filename", "")).lower().endswith((".xlsx", ".xlsm")):
        raise HoroshopPricesError("Підтримуються лише Excel-файли .xlsx та .xlsm.")
    contents = await uploaded.read()
    if not contents or len(contents) > MAX_UPLOAD_BYTES:
        raise HoroshopPricesError("Excel-файл порожній або перевищує 20 МБ.")
    mappings = stored_mappings()
    additional_mappings = await uploaded_mappings(form)
    if additional_mappings is not None:
        mappings = await asyncio.to_thread(add_to_stored_mappings, additional_mappings)
    return parse_excel_prices(contents), credentials_from(dict(form)), parse_bool(form.get("normalize_articles"), True), mappings


app = FastAPI(title="Ціни Хорошоп")


@app.exception_handler(Exception)
async def unexpected_error(_: Request, error: Exception) -> JSONResponse:
    logger.exception("Unhandled web request error", exc_info=error)
    return JSONResponse(status_code=500, content={"detail": "Внутрішня помилка сервера. Перевірте публічний лог."})


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((PROJECT_DIR / "web_ui.html").read_text(encoding="utf-8"), headers={"Cache-Control": "no-store"})


@app.get("/api/template")
def template() -> Response:
    return Response(build_excel_template(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": 'attachment; filename="horoshop_prices_template.xlsx"', "Cache-Control": "no-store"})


@app.post("/api/mappings")
async def upload_mappings(request: Request) -> dict[str, int]:
    try:
        mappings = await uploaded_mappings(await request.form())
        if mappings is None:
            raise HoroshopPricesError("Оберіть Excel-файл сопоставлень.")
        stored = await asyncio.to_thread(add_to_stored_mappings, mappings)
        return {"sources": len(stored.entries), "rules": sum(len(targets) for targets in stored.entries.values())}
    except (HoroshopPricesError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/preview")
async def preview(request: Request) -> dict[str, Any]:
    try:
        rows, credentials, normalize_articles, mappings = await uploaded_rows(request)
        return await asyncio.to_thread(preview_rows, rows, credentials, normalize_articles, mappings)
    except (HoroshopPricesError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/import")
async def import_prices(request: Request) -> dict[str, Any]:
    try:
        rows, credentials, normalize_articles, mappings = await uploaded_rows(request)
        return await asyncio.to_thread(execute_rows, rows, credentials, normalize_articles, mappings)
    except (HoroshopPricesError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/price")
async def one_price(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        price_type = parse_price_type(data.get("price_type"))
        raw_value = normalize(data.get("value"))
        raw_percent = normalize(data.get("percent"))
        if raw_value and raw_percent:
            raise ValueError("Вкажіть суму в грн або відсоток від РРЦ, але не обидва значення.")
        delete = raw_value.casefold() in {"видалити", "delete"}
        if delete and price_type == "price":
            raise ValueError("Поточну ціну не можна видалити. Вкажіть нове значення.")
        value = None if delete else (parse_decimal(raw_value) if raw_value else None)
        percent = parse_decimal(raw_percent, "Відсоток від РРЦ") if raw_percent else None
        if percent is not None and not Decimal("1") <= percent <= Decimal("100"):
            raise ValueError("Відсоток від РРЦ має бути від 1 до 100.")
        if value is None and percent is None and not delete:
            raise ValueError("Вкажіть суму в грн, відсоток від РРЦ або «Видалити».")
        if percent is not None and price_type != "price":
            raise ValueError("Відсоток від РРЦ доступний лише для поточної ціни.")
        threshold = parse_threshold(data.get("threshold"))
        current = FieldChange(value=value) if price_type == "price" else FieldChange()
        old = FieldChange(value=value, delete=delete) if price_type == "price_old" else FieldChange()
        wholesale = () if not price_type.startswith("wholesale_") else (WholesaleChange(int(price_type.rsplit("_", 1)[1]), FieldChange(value=value, delete=delete), threshold),)
        row = PriceRow(normalize(data.get("article")), current, percent, old, False, False, wholesale, 1)
        return await asyncio.to_thread(execute_rows, [row], credentials_from(data), parse_bool(data.get("normalize_articles"), True), stored_mappings())
    except (HoroshopPricesError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


def run_server() -> None:
    import uvicorn
    runtime_settings = get_settings()
    configure_service_output(runtime_settings)
    uvicorn.run(app, host=runtime_settings.host, port=runtime_settings.port)


if __name__ == "__main__":
    run_server()
