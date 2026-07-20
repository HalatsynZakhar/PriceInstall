from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


DEFAULT_EXPORT_LIMIT = 500
PUBLIC_LOG_PATH_JSON_PATTERN = re.compile(r'("public_log_path"\s*:\s*")(?P<value>[^"]*)(")')
PRICE_TYPES = {
    "поточна ціна": "price", "ціна": "price", "price": "price", "current": "price", "акційна ціна": "price",
    "ррц": "price_old", "стара ціна": "price_old", "old price": "price_old", "price_old": "price_old",
    "опт 1": "wholesale_1", "оптова ціна 1": "wholesale_1", "wholesale 1": "wholesale_1",
    "опт 2": "wholesale_2", "оптова ціна 2": "wholesale_2", "wholesale 2": "wholesale_2",
    "опт 3": "wholesale_3", "оптова ціна 3": "wholesale_3", "wholesale 3": "wholesale_3",
    "опт 4": "wholesale_4", "оптова ціна 4": "wholesale_4", "wholesale 4": "wholesale_4",
    "опт 5": "wholesale_5", "оптова ціна 5": "wholesale_5", "wholesale 5": "wholesale_5",
}


class HoroshopPricesError(RuntimeError):
    pass


@dataclass(frozen=True)
class WholesaleRule:
    tier: int
    minimal_threshold: int
    discount_percent: Decimal


@dataclass(frozen=True)
class Settings:
    domain: str
    host: str
    port: int
    batch_size: int
    request_timeout_seconds: int
    public_log_path: Path
    public_log_name: str
    wholesale_rules: tuple[WholesaleRule, ...]
    recalculate_wholesale_on_current_price: bool

    @property
    def public_log_file(self) -> Path:
        return self.public_log_path / self.public_log_name


@dataclass(frozen=True)
class Credentials:
    login: str = ""
    password: str = ""
    token: str = ""


@dataclass(frozen=True)
class PriceRow:
    display_article: str
    price_type: str
    value: Decimal
    threshold: int | None
    row_number: int


@dataclass(frozen=True)
class CatalogProduct:
    article: str
    display_article: str
    price: Decimal | None
    price_old: Decimal | None
    wholesale_prices: tuple[dict[str, Any], ...]


@dataclass(frozen=True)
class PricePlan:
    article: str
    display_article: str
    payload: dict[str, Any]
    rows: tuple[PriceRow, ...]
    warnings: tuple[str, ...] = ()
    error: str = ""

    @property
    def ready(self) -> bool:
        return not self.error


def normalize(value: Any) -> str:
    return "" if value is None else str(value).strip()


def endpoint_url(domain: str, endpoint: str) -> str:
    return urljoin(f"{domain.rstrip('/')}/", endpoint.lstrip("/"))


def parse_decimal(value: Any, field: str = "ціна", allow_zero: bool = False) -> Decimal:
    text = normalize(value).replace(" ", "").replace(",", ".")
    try:
        number = Decimal(text)
    except (InvalidOperation, ValueError) as error:
        raise ValueError(f"{field} має бути числом, більшим за нуль.") from error
    if number < 0 or (number == 0 and not allow_zero):
        raise ValueError(f"{field} має бути більшим за нуль.")
    return number.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def parse_threshold(value: Any) -> int | None:
    text = normalize(value)
    if not text:
        return None
    try:
        number = int(text)
    except ValueError as error:
        raise ValueError("кількість від має бути цілим числом.") from error
    if number < 1:
        raise ValueError("кількість від має бути не меншою за 1.")
    return number


def parse_price_type(value: Any) -> str:
    key = normalize(value).casefold()
    price_type = PRICE_TYPES.get(key)
    if not price_type:
        raise ValueError("невідомий тип ціни. Використайте: Поточна ціна, РРЦ або Опт 1-5.")
    return price_type


def _repair_path(config_text: str) -> str:
    def replace(match: re.Match[str]) -> str:
        value = re.sub(r"(?<!\\)\\(?!\\)", lambda _: "\\\\", match.group("value"))
        return f"{match.group(1)}{value}{match.group(3)}"
    return PUBLIC_LOG_PATH_JSON_PATTERN.sub(replace, config_text)


def load_settings(config_file: Path) -> Settings:
    config_text = config_file.read_text(encoding="utf-8-sig")
    try:
        raw = json.loads(config_text)
    except json.JSONDecodeError as original_error:
        repaired = _repair_path(config_text)
        if repaired == config_text:
            raise original_error
        raw = json.loads(repaired)
        config_file.write_text(repaired, encoding="utf-8")
    if not isinstance(raw, dict):
        raise ValueError("config.json має містити об'єкт.")
    server = raw.get("server") or {}
    horoshop = raw.get("horoshop") or {}
    logging_config = raw.get("logging") or {}
    prices = raw.get("wholesale_defaults") or {}
    domain = normalize(horoshop.get("domain"))
    if not domain:
        raise ValueError("Вкажіть horoshop.domain у config.json.")
    rules: list[WholesaleRule] = []
    raw_rules = prices.get("rules") if isinstance(prices, dict) else None
    if raw_rules is None:
        raw_rules = [
            {"tier": 1, "minimal_threshold": 2, "discount_percent": 10},
            {"tier": 2, "minimal_threshold": 5, "discount_percent": 15},
            {"tier": 3, "minimal_threshold": 10, "discount_percent": 25},
        ]
    if not isinstance(raw_rules, list):
        raise ValueError("wholesale_defaults.rules має бути списком.")
    for raw_rule in raw_rules:
        if not isinstance(raw_rule, dict):
            raise ValueError("Кожне правило оптової ціни має бути об'єктом.")
        tier = int(raw_rule.get("tier", 0))
        threshold = int(raw_rule.get("minimal_threshold", 0))
        discount = parse_decimal(raw_rule.get("discount_percent"), "відсоток знижки")
        if tier not in range(1, 6) or threshold < 1 or discount >= 100:
            raise ValueError("Правила опту повинні містити рівень 1-5, поріг від 1 та знижку до 100%.")
        rules.append(WholesaleRule(tier, threshold, discount))
    path = Path(normalize(logging_config.get("public_log_path", "logs")) or "logs")
    if not path.is_absolute():
        path = config_file.parent / path
    name = normalize(logging_config.get("public_log_name", "horoshop_prices.log")) or "horoshop_prices.log"
    if Path(name).name != name:
        raise ValueError("logging.public_log_name має бути лише назвою файлу.")
    return Settings(
        domain=domain.rstrip("/"), host=normalize(server.get("host", "0.0.0.0")) or "0.0.0.0",
        port=max(1, min(65535, int(server.get("port", 8095)))), batch_size=max(1, int(horoshop.get("batch_size", 50))),
        request_timeout_seconds=max(1, int(horoshop.get("request_timeout_seconds", 60))), public_log_path=path, public_log_name=name,
        wholesale_rules=tuple(sorted(rules, key=lambda rule: rule.tier)),
        recalculate_wholesale_on_current_price=bool(prices.get("recalculate_on_current_price", True)) if isinstance(prices, dict) else True,
    )


def parse_excel_prices(data: bytes) -> list[PriceRow]:
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        rows: list[PriceRow] = []
        headers = {"артикул", "артикул для відображення", "article", "article_for_display"}
        for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
            if not row or all(value is None for value in row[:4]):
                continue
            article = normalize(row[0] if len(row) > 0 else "")
            if article.casefold() in headers:
                continue
            try:
                if not article:
                    raise ValueError("вкажіть артикул.")
                price_type = parse_price_type(row[1] if len(row) > 1 else "")
                value = parse_decimal(row[2] if len(row) > 2 else "", allow_zero=price_type == "price_old")
                threshold = parse_threshold(row[3] if len(row) > 3 else "")
                if not price_type.startswith("wholesale_") and threshold is not None:
                    raise ValueError("кількість від заповнюється тільки для оптової ціни.")
            except ValueError as error:
                raise ValueError(f"Рядок {row_number}: {error}") from error
            rows.append(PriceRow(article, price_type, value, threshold, row_number))
        if not rows:
            raise ValueError("Excel не містить жодної ціни.")
        return rows
    finally:
        workbook.close()


def _style_header(worksheet: Any) -> None:
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="6D3B8F")
        cell.alignment = Alignment(horizontal="center")


def build_excel_template() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Ціни"
    sheet.append(["Артикул для відображення", "Тип ціни", "Значення", "Кількість від (лише для Опт 4-5)"])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = "A1:D1"
    for column, width in {"A": 30, "B": 24, "C": 16, "D": 32}.items():
        sheet.column_dimensions[column].width = width
    _style_header(sheet)
    examples = [("01063", "Поточна ціна", 1000, ""), ("01063", "РРЦ", 1200, ""), ("01063", "Опт 4", 700, 20)]
    for row in examples:
        sheet.append(row)
    for row in range(2, 202):
        sheet.cell(row=row, column=1).number_format = "@"
    guide = workbook.create_sheet("Інструкція")
    guide.column_dimensions["A"].width = 110
    guide["A1"] = "Масове встановлення цін Хорошоп"
    guide["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    guide["A1"].fill = PatternFill("solid", fgColor="6D3B8F")
    guide.merge_cells("A1:B1")
    notes = [
        "Кожен рядок оновлює один тип ціни одного товару. Один товар може мати кілька рядків.",
        "Артикул можна вказати як article_for_display зі сайту або внутрішній article Хорошоп. Пошук: точний, потім без урахування регістру.",
        "Типи: Поточна ціна, РРЦ, Опт 1, Опт 2, Опт 3, Опт 4, Опт 5.",
        "Для поточної ціни за замовчуванням опт 1-3 перераховується за правилами 2 шт. -10%, 5 шт. -15%, 10 шт. -25%. Правила можна змінити у config.json.",
        "Оптові ціни з ціною не нижче поточної автоматично виключаються. Інші наявні рівні опта зберігаються.",
        "Для Опт 4-5 заповніть «Кількість від». Для Опт 1-3 поріг береться з config.json, якщо колонка порожня.",
    ]
    for index, note in enumerate(notes, start=3):
        guide.cell(index, 1, note).alignment = Alignment(wrap_text=True, vertical="top")
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class CatalogIndex:
    def __init__(self, products: list[CatalogProduct]) -> None:
        self.products = products
        self.by_article = {item.article: item for item in products}
        self.by_exact: dict[str, list[CatalogProduct]] = {}
        self.by_folded: dict[str, list[CatalogProduct]] = {}
        for item in products:
            for key in {item.article, item.display_article} - {""}:
                self.by_exact.setdefault(key, []).append(item)
                self.by_folded.setdefault(key.casefold(), []).append(item)

    @classmethod
    def from_raw(cls, raw_products: list[dict[str, Any]]) -> "CatalogIndex":
        products: list[CatalogProduct] = []
        for raw in raw_products:
            article = normalize(raw.get("article")) if isinstance(raw, dict) else ""
            if not article:
                continue
            wholesale = raw.get("wholesale_prices", [])
            products.append(CatalogProduct(
                article, normalize(raw.get("article_for_display")), decimal_or_none(raw.get("price")), decimal_or_none(raw.get("price_old")),
                tuple(item for item in wholesale if isinstance(item, dict)) if isinstance(wholesale, list) else (),
            ))
        return cls(products)

    def resolve(self, value: str) -> tuple[CatalogProduct | None, str]:
        exact = list({item.article: item for item in self.by_exact.get(value, [])}.values())
        if len(exact) == 1:
            return exact[0], ""
        if len(exact) > 1:
            return None, f"Артикул '{value}' не є унікальним."
        folded = list({item.article: item for item in self.by_folded.get(value.casefold(), [])}.values())
        if len(folded) == 1:
            return folded[0], ""
        if len(folded) > 1:
            return None, f"Артикул '{value}' не є унікальним."
        return None, f"Артикул '{value}' не знайдений у каталозі."


def decimal_or_none(value: Any) -> Decimal | None:
    text = normalize(value)
    if not text or text in {"0", "0.0", "0.00"}:
        return None
    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return None


def rule_for(settings: Settings, tier: int) -> WholesaleRule | None:
    return next((rule for rule in settings.wholesale_rules if rule.tier == tier), None)


def wholesale_entries(existing: tuple[dict[str, Any], ...]) -> dict[int, Decimal]:
    entries: dict[int, Decimal] = {}
    for item in existing:
        try:
            threshold = int(item.get("minimal_threshold"))
            price = parse_decimal(item.get("price"))
        except (TypeError, ValueError):
            continue
        entries[threshold] = price
    return entries


def plan_prices(rows: list[PriceRow], catalog: CatalogIndex, settings: Settings) -> list[PricePlan]:
    grouped: dict[str, list[PriceRow]] = {}
    errors: list[PricePlan] = []
    for row in rows:
        product, error = catalog.resolve(row.display_article)
        if error or product is None:
            errors.append(PricePlan("", row.display_article, {}, (row,), error=error))
        else:
            grouped.setdefault(product.article, []).append(row)
    plans = errors
    for article, article_rows in grouped.items():
        product = catalog.by_article[article]
        current = product.price
        old = product.price_old
        direct_wholesale = any(row.price_type.startswith("wholesale_") for row in article_rows)
        has_current = any(row.price_type == "price" for row in article_rows)
        for row in article_rows:
            if row.price_type == "price": current = row.value
            elif row.price_type == "price_old": old = row.value
        if current is None:
            plans.append(PricePlan(article, product.display_article or article, {}, tuple(article_rows), error="Немає поточної ціни: спочатку встановіть «Поточна ціна»."))
            continue
        wholesale = wholesale_entries(product.wholesale_prices)
        warnings: list[str] = []
        if has_current and settings.recalculate_wholesale_on_current_price:
            for rule in settings.wholesale_rules:
                if rule.tier <= 3:
                    wholesale[rule.minimal_threshold] = (current * (Decimal("100") - rule.discount_percent) / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            warnings.append("Опт 1-3 перераховано за правилами за замовчуванням.")
        for row in article_rows:
            if not row.price_type.startswith("wholesale_"):
                continue
            tier = int(row.price_type.rsplit("_", 1)[1])
            rule = rule_for(settings, tier)
            threshold = row.threshold or (rule.minimal_threshold if rule else None)
            if threshold is None:
                plans.append(PricePlan(article, product.display_article or article, {}, tuple(article_rows), error=f"Для Опт {tier} вкажіть «Кількість від" + "."))
                break
            wholesale[threshold] = row.value
        else:
            invalid = [threshold for threshold, value in wholesale.items() if value >= current]
            for threshold in invalid:
                del wholesale[threshold]
            if invalid:
                warnings.append("Вимкнено оптові рівні, де ціна не нижча за поточну: " + ", ".join(map(str, sorted(invalid))) + ".")
            payload: dict[str, Any] = {"article": article}
            if any(row.price_type == "price" for row in article_rows): payload["price"] = float(current)
            if any(row.price_type == "price_old" for row in article_rows): payload["price_old"] = float(old) if old else 0
            if has_current or direct_wholesale:
                payload["wholesale_prices"] = [{"minimal_threshold": threshold, "price": float(value)} for threshold, value in sorted(wholesale.items())]
            plans.append(PricePlan(article, product.display_article or article, payload, tuple(article_rows), tuple(warnings)))
    return sorted(plans, key=lambda plan: min(row.row_number for row in plan.rows))


class HoroshopClient:
    def __init__(self, settings: Settings, credentials: Credentials) -> None:
        self.settings, self.credentials, self.session, self._token = settings, credentials, requests.Session(), credentials.token

    def _post(self, endpoint: str, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            response = self.session.post(endpoint_url(self.settings.domain, endpoint), json=payload, timeout=self.settings.request_timeout_seconds)
            response.raise_for_status()
            data = response.json()
        except requests.RequestException as error:
            raise HoroshopPricesError(f"Помилка запиту до Хорошоп: {error}") from error
        except ValueError as error:
            raise HoroshopPricesError("Хорошоп повернув відповідь не у форматі JSON.") from error
        if not isinstance(data, dict):
            raise HoroshopPricesError("Хорошоп повернув некоректну відповідь.")
        if str(data.get("status", "")).upper() in {"ERROR", "EXCEPTION"}:
            raise HoroshopPricesError(str(data))
        return data

    def token(self) -> str:
        if self._token: return self._token
        if not self.credentials.login or not self.credentials.password:
            raise HoroshopPricesError("Вкажіть логін і пароль API або чинний токен.")
        response = self._post("/api/auth/", {"login": self.credentials.login, "password": self.credentials.password})
        token = response.get("response", {}).get("token")
        if not token: raise HoroshopPricesError("Хорошоп не повернув токен авторизації.")
        self._token = str(token)
        return self._token

    def export_catalog(self) -> list[dict[str, Any]]:
        offset, products = 0, []
        while True:
            response = self._post("/api/catalog/export/", {"token": self.token(), "offset": offset, "limit": DEFAULT_EXPORT_LIMIT, "includedParams": ["article_for_display", "price", "price_old", "wholesale_prices"]})
            nested = response.get("response")
            page = nested.get("products") if isinstance(nested, dict) else response.get("products")
            if not isinstance(page, list): raise HoroshopPricesError("Експорт каталогу не містить товарів.")
            products.extend(item for item in page if isinstance(item, dict))
            if len(page) < DEFAULT_EXPORT_LIMIT: return products
            offset += DEFAULT_EXPORT_LIMIT

    def import_products(self, products: list[dict[str, Any]]) -> dict[str, Any]:
        if not products: return {"status": "OK", "response": {"log": []}}
        return self._post("/api/catalog/import/", {"token": self.token(), "products": products})


def import_results(response: dict[str, Any]) -> dict[str, tuple[bool, str]]:
    status = str(response.get("status", "")).upper()
    nested = response.get("response")
    entries = nested.get("log", []) if isinstance(nested, dict) else []
    results: dict[str, tuple[bool, str]] = {}
    for entry in entries if isinstance(entries, list) else []:
        if not isinstance(entry, dict): continue
        article = normalize(entry.get("article"))
        info = entry.get("info", [])
        codes, messages = [], []
        for item in info if isinstance(info, list) else []:
            if isinstance(item, dict): codes.append(item.get("code")); messages.append(normalize(item.get("message")))
        results[article] = (0 in codes or (status == "OK" and not codes), "; ".join(message for message in messages if message))
    return results
