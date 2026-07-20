from __future__ import annotations

import asyncio
import logging
import sys
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from horoshop_prices import (
    CatalogIndex, Credentials, HoroshopClient, HoroshopPricesError, PricePlan, PriceRow, Settings,
    build_excel_template, import_results, load_settings, normalize, parse_decimal, parse_excel_prices, parse_price_type, parse_threshold, plan_prices,
)


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
settings: Settings | None = None
logger = logging.getLogger(__name__)


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
    return {
        "article": plan.display_article,
        "internal_article": plan.article,
        "rows": [{"row_number": row.row_number, "price_type": row.price_type, "value": str(row.value), "threshold": row.threshold} for row in plan.rows],
        "status": status or ("ready" if plan.ready else "invalid"),
        "message": message or plan.error or "; ".join(plan.warnings),
    }


def preview_rows(rows: list[PriceRow], credentials: Credentials) -> dict[str, Any]:
    catalog, _ = catalog_for(credentials)
    plans = plan_prices(rows, catalog, get_settings())
    return {"items": [serialise(plan) for plan in plans], "ready": sum(plan.ready for plan in plans), "errors": sum(not plan.ready for plan in plans)}


def execute_rows(rows: list[PriceRow], credentials: Credentials) -> dict[str, Any]:
    catalog, client = catalog_for(credentials)
    plans = plan_prices(rows, catalog, get_settings())
    payload = [plan.payload for plan in plans if plan.ready]
    api_results: dict[str, tuple[bool, str]] = {}
    for start in range(0, len(payload), get_settings().batch_size):
        api_results.update(import_results(client.import_products(payload[start:start + get_settings().batch_size])))
    items: list[dict[str, Any]] = []
    for plan in plans:
        if not plan.ready:
            items.append(serialise(plan, "invalid"))
            continue
        success, message = api_results.get(plan.article, (False, "API не повернуло результат для товару."))
        joined = "; ".join(item for item in ("; ".join(plan.warnings), message) if item)
        items.append(serialise(plan, "synced" if success else "error", joined))
    return {"items": items, "imported": sum(item["status"] == "synced" for item in items), "errors": sum(item["status"] in {"error", "invalid"} for item in items)}


async def uploaded_rows(request: Request) -> tuple[list[PriceRow], Credentials]:
    form = await request.form()
    uploaded = form.get("file")
    if uploaded is None or not hasattr(uploaded, "read"):
        raise HoroshopPricesError("Оберіть Excel-файл .xlsx або .xlsm.")
    if not str(getattr(uploaded, "filename", "")).lower().endswith((".xlsx", ".xlsm")):
        raise HoroshopPricesError("Підтримуються лише Excel-файли .xlsx та .xlsm.")
    contents = await uploaded.read()
    if not contents or len(contents) > MAX_UPLOAD_BYTES:
        raise HoroshopPricesError("Excel-файл порожній або перевищує 20 МБ.")
    return parse_excel_prices(contents), credentials_from(dict(form))


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


@app.post("/api/preview")
async def preview(request: Request) -> dict[str, Any]:
    try:
        rows, credentials = await uploaded_rows(request)
        return await asyncio.to_thread(preview_rows, rows, credentials)
    except (HoroshopPricesError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/import")
async def import_prices(request: Request) -> dict[str, Any]:
    try:
        rows, credentials = await uploaded_rows(request)
        return await asyncio.to_thread(execute_rows, rows, credentials)
    except (HoroshopPricesError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


@app.post("/api/price")
async def one_price(request: Request) -> dict[str, Any]:
    try:
        data = await request.json()
        price_type = parse_price_type(data.get("price_type"))
        row = PriceRow(normalize(data.get("article")), price_type, parse_decimal(data.get("value"), allow_zero=price_type == "price_old"), parse_threshold(data.get("threshold")), 1)
        return await asyncio.to_thread(execute_rows, [row], credentials_from(data))
    except (HoroshopPricesError, ValueError) as error:
        raise HTTPException(status_code=400, detail=str(error)) from error


def run_server() -> None:
    import uvicorn
    runtime_settings = get_settings()
    configure_service_output(runtime_settings)
    uvicorn.run(app, host=runtime_settings.host, port=runtime_settings.port)


if __name__ == "__main__":
    run_server()
