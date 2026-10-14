from __future__ import annotations

import io
import json
import re
from itertools import chain
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_DOWN, ROUND_HALF_UP
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
    recalculate_wholesale_from_rrc: bool

    @property
    def public_log_file(self) -> Path:
        return self.public_log_path / self.public_log_name


@dataclass(frozen=True)
class Credentials:
    login: str = ""
    password: str = ""
    token: str = ""


@dataclass(frozen=True)
class FieldChange:
    value: Decimal | None = None
    delete: bool = False

    @property
    def specified(self) -> bool:
        return self.value is not None or self.delete


@dataclass(frozen=True)
class WholesaleChange:
    tier: int
    change: FieldChange
    threshold: int | None = None


@dataclass(frozen=True)
class PriceRow:
    display_article: str
    current: FieldChange
    current_percent_of_old: Decimal | None
    old: FieldChange
    move_current_to_old: bool
    move_old_to_current: bool
    wholesale: tuple[WholesaleChange, ...]
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
    catalog_product: CatalogProduct | None = None

    @property
    def ready(self) -> bool:
        return not self.error


def normalize(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def leading_zeros_count(text: str) -> int:
    count = 0
    for character in text:
        if character == "0":
            count += 1
        else:
            break
    return count


def normalize_article_code(value: Any) -> str:
    text = normalize(value)
    if not text:
        return ""
    raw_alpha = "".join(character if character.isalnum() else "-" for character in text).strip("-")
    zeros = leading_zeros_count(raw_alpha)
    cleaned = re.sub(r"-{2,}", "-", raw_alpha)
    if zeros > 0 and not cleaned.startswith("0" * zeros):
        cleaned = ("0" * zeros) + cleaned.lstrip("0")
    return cleaned


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
        recalculate_wholesale_from_rrc=bool(prices.get("recalculate_from_rrc", prices.get("recalculate_on_current_price", True))) if isinstance(prices, dict) else True,
    )


DELETE_MARKERS = {"видалити", "delete", "так", "yes"}
YES_MARKERS = {"так", "да", "yes", "1"}
HEADER_ARTICLE = {"артикул", "артикул для відображення", "article", "article_for_display"}
MAPPING_HEADER_MARKERS = ("артикул", "article", "варіант", "вариант", "variant", "подартикул", "підартикул", "субартикул", "source", "target", "ключ")
@dataclass(frozen=True)
class ArticleMappings:
    entries: dict[str, tuple[str, ...]]

    @classmethod
    def empty(cls) -> "ArticleMappings":
        return cls({})

    def merged_with(self, other: "ArticleMappings") -> "ArticleMappings":
        combined = {source: list(targets) for source, targets in self.entries.items()}
        for source, targets in other.entries.items():
            combined.setdefault(source, []).extend(targets)
        result: dict[str, tuple[str, ...]] = {}
        for source, targets in combined.items():
            deduped = tuple(dict.fromkeys(targets))
            if any(target.casefold() != source.casefold() for target in deduped):
                result[source] = deduped
        return ArticleMappings(result)

    def targets_for(self, article: str, use_normalized: bool) -> tuple[str, ...]:
        if not self.entries:
            return (article,)
        exact = self.entries.get(article)
        if exact:
            return exact
        folded = article.casefold()
        folded_matches = [targets for source, targets in self.entries.items() if source.casefold() == folded]
        if len(folded_matches) == 1:
            return folded_matches[0]
        if use_normalized:
            norm_article = normalize_article_code(article)
            zeros = leading_zeros_count(norm_article)
            normalized = norm_article.casefold()
            normalized_matches = [
                targets for source, targets in self.entries.items()
                if normalize_article_code(source).casefold() == normalized
                and leading_zeros_count(normalize_article_code(source)) == zeros
            ]
            if len(normalized_matches) == 1:
                return normalized_matches[0]
        return (article,)


def header_key(value: Any) -> str:
    return " ".join(normalize(value).casefold().split())


def parse_change(value: Any, label: str) -> FieldChange:
    text = normalize(value)
    if not text:
        return FieldChange()
    if text.casefold() in DELETE_MARKERS:
        return FieldChange(delete=True)
    return FieldChange(value=parse_decimal(text, label))


def parse_yes(value: Any, label: str) -> bool:
    text = normalize(value)
    if not text or text.casefold() in {"ні", "нет", "no", "0"}:
        return False
    if text.casefold() in YES_MARKERS:
        return True
    raise ValueError(f"у колонці «{label}» вкажіть Так або залиште порожньою.")


def column_map(headers: tuple[Any, ...]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    aliases = {"поточна ціна": "price", "поточна ціна (грн)": "price", "поточна ціна, % від ррц": "percent", "поточна ціна (% від ррц)": "percent", "ррц": "old", "ррц (грн)": "old", "поточну ціну в ррц (так)": "move", "ррц в поточну ціну (так)": "restore"}
    for tier in range(1, 6):
        aliases[f"опт {tier}"] = f"wholesale_{tier}"
        aliases[f"опт {tier} (грн)"] = f"wholesale_{tier}"
        aliases[f"опт {tier} від"] = f"threshold_{tier}"
        aliases[f"опт {tier} від (шт.)"] = f"threshold_{tier}"
    for index, value in enumerate(headers):
        key = header_key(value)
        target = "article" if key in HEADER_ARTICLE else aliases.get(key)
        if target:
            if target in mapping:
                raise ValueError(f"Заголовок «{value}» повторюється.")
            mapping[target] = index
    if "article" not in mapping:
        raise ValueError("Не знайдено обов'язковий стовпець «Артикул».")
    return mapping


def cell(row: tuple[Any, ...], mapping: dict[str, int], key: str) -> Any:
    index = mapping.get(key)
    return row[index] if index is not None and index < len(row) else None


def parse_excel_prices(data: bytes) -> list[PriceRow]:
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        iterator = worksheet.iter_rows(values_only=True)
        try:
            headers = next(iterator)
        except StopIteration as error:
            raise ValueError("Excel-файл порожній.") from error
        mapping = column_map(headers)
        rows: list[PriceRow] = []
        for row_number, row in enumerate(iterator, start=2):
            if not row or all(value is None for value in row):
                continue
            try:
                article = normalize(cell(row, mapping, "article"))
                if not article:
                    raise ValueError("вкажіть артикул.")
                current = parse_change(cell(row, mapping, "price"), "Поточна ціна")
                if current.delete:
                    raise ValueError("поточну ціну не можна видалити. Вкажіть нове значення.")
                percent_text = normalize(cell(row, mapping, "percent"))
                percent = parse_decimal(percent_text, "Відсоток від РРЦ") if percent_text else None
                if percent is not None and not Decimal("1") <= percent <= Decimal("100"):
                    raise ValueError("відсоток від РРЦ має бути від 1 до 100.")
                if current.value is not None and percent is not None:
                    raise ValueError("вкажіть або поточну ціну, або відсоток від РРЦ, але не обидва значення.")
                old = parse_change(cell(row, mapping, "old"), "РРЦ")
                move = parse_yes(cell(row, mapping, "move"), "Поточну ціну в РРЦ (Так)")
                restore = parse_yes(cell(row, mapping, "restore"), "РРЦ в поточну ціну (Так)")
                if move and restore:
                    raise ValueError("в одному рядку можна обрати лише один автоматичний перенос ціни.")
                if move and old.specified:
                    raise ValueError("не поєднуйте РРЦ та «Поточну ціну в РРЦ (Так)» в одному рядку.")
                if move and current.value is None and percent is None:
                    raise ValueError("для перенесення поточної ціни в РРЦ вкажіть нову поточну ціну або відсоток від РРЦ.")
                if restore and old.value is not None:
                    raise ValueError("для перенесення РРЦ у поточну ціну не вказуйте нове значення РРЦ.")
                if restore and percent is not None:
                    raise ValueError("не поєднуйте перенесення РРЦ у поточну ціну з відсотком від РРЦ.")
                if percent is not None and old.delete:
                    raise ValueError("не можна видаляти РРЦ, з якого обчислюється поточна ціна.")
                wholesale: list[WholesaleChange] = []
                for tier in range(1, 6):
                    change = parse_change(cell(row, mapping, f"wholesale_{tier}"), f"Опт {tier}")
                    threshold = parse_threshold(cell(row, mapping, f"threshold_{tier}"))
                    if threshold is not None and not change.specified:
                        raise ValueError(f"для Опт {tier} задано кількість без ціни або команди «Видалити».")
                    if change.specified:
                        wholesale.append(WholesaleChange(tier, change, threshold))
                if not current.specified and percent is None and not old.specified and not move and not restore and not wholesale:
                    continue
                rows.append(PriceRow(article, current, percent, old, move, restore, tuple(wholesale), row_number))
            except ValueError as error:
                raise ValueError(f"Рядок {row_number}: {error}") from error
        if not rows:
            raise ValueError("У файлі немає жодної команди. Порожні клітинки не змінюють ціни.")
        return rows
    finally:
        workbook.close()


def mapping_column_indexes(headers: tuple[Any, ...]) -> tuple[int, ...]:
    indexes = tuple(index for index, value in enumerate(headers) if normalize(value))
    if len(indexes) < 2:
        raise ValueError("У файлі сопоставлень має бути мінімум два стовпці з варіантами артикула.")
    return indexes


def is_mapping_header(row: tuple[Any, ...]) -> bool:
    return any(marker in header_key(value) for value in row for marker in MAPPING_HEADER_MARKERS)


def parse_article_mappings(data: bytes) -> ArticleMappings:
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        iterator = worksheet.iter_rows(values_only=True)
        try:
            first_row = next(iterator)
        except StopIteration as error:
            raise ValueError("Файл сопоставлень порожній.") from error
        has_header = is_mapping_header(first_row)
        if has_header:
            column_indexes = mapping_column_indexes(first_row)
            rows = iterator
            start_row = 2
        else:
            column_indexes = tuple(range(len(first_row)))
            rows = chain((first_row,), iterator)
            start_row = 1
        mappings: dict[str, list[str]] = {}
        for row_number, row in enumerate(rows, start=start_row):
            if not row or all(value is None for value in row):
                continue
            if has_header:
                raw_articles = [
                    normalize(row[index] if index < len(row) else None)
                    for index in column_indexes
                    if normalize(row[index] if index < len(row) else None)
                ]
            else:
                raw_articles = [normalize(value) for value in row if normalize(value)]
            unique_articles: list[str] = list(dict.fromkeys(raw_articles))
            if len(unique_articles) < 2:
                continue
            if len({a.casefold() for a in unique_articles}) < 2:
                continue
            for i in range(len(unique_articles)):
                for j in range(i + 1, len(unique_articles)):
                    a, b = unique_articles[i], unique_articles[j]
                    if a.lstrip("0") == b.lstrip("0") and a != b:
                        raise ValueError(
                            f"Рядок {row_number}: артикули «{a}» та «{b}» відрізняються лише початковими нулями (це різні бренди/товари). "
                            f"Їх заборонено зіставляти один з одним."
                        )
            for source in unique_articles:
                mappings.setdefault(source, []).extend(unique_articles)

        final_entries: dict[str, tuple[str, ...]] = {}
        for source, targets in mappings.items():
            deduped = tuple(dict.fromkeys(targets))
            if any(target.casefold() != source.casefold() for target in deduped):
                final_entries[source] = deduped

        if not final_entries:
            raise ValueError("У файлі сопоставлень немає жодного корисного правила (усі рядки порожні, містять лише один артикул або однакові значення).")
        return ArticleMappings(final_entries)
    finally:
        workbook.close()


def load_article_mappings(path: Path) -> ArticleMappings:
    if not path.exists():
        return ArticleMappings.empty()
    try:
        raw_entries = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise HoroshopPricesError(f"Не вдалося прочитати базу сопоставлень: {path}.") from error
    if not isinstance(raw_entries, dict):
        raise HoroshopPricesError("База сопоставлень має некоректний формат.")
    entries: dict[str, tuple[str, ...]] = {}
    for source, targets in raw_entries.items():
        source_article = normalize(source)
        if not source_article or not isinstance(targets, list):
            raise HoroshopPricesError("База сопоставлень має некоректний формат.")
        target_articles = tuple(dict.fromkeys(normalize(target) for target in targets if normalize(target)))
        if not target_articles:
            raise HoroshopPricesError("База сопоставлень має некоректний формат.")
        if any(target.casefold() != source_article.casefold() for target in target_articles):
            entries[source_article] = target_articles
    return ArticleMappings(entries)


def save_article_mappings(path: Path, mappings: ArticleMappings) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = {source: list(targets) for source, targets in mappings.entries.items()}
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    try:
        temporary_path.write_text(json.dumps(serialized, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary_path.replace(path)
    except OSError as error:
        raise HoroshopPricesError(f"Не вдалося зберегти базу сопоставлень: {path}.") from error


def price_row_excel_values(row: PriceRow) -> list[Any]:
    values: list[Any] = [
        row.row_number,
        row.display_article,
        float(row.current.value) if row.current.value is not None else "",
        float(row.current_percent_of_old) if row.current_percent_of_old is not None else "",
        "Видалити" if row.old.delete else (float(row.old.value) if row.old.value is not None else ""),
        "Так" if row.move_current_to_old else "",
        "Так" if row.move_old_to_current else "",
    ]
    by_tier = {item.tier: item for item in row.wholesale}
    for tier in range(1, 6):
        item = by_tier.get(tier)
        if item is None:
            values.extend(["", ""])
        else:
            values.extend(["Видалити" if item.change.delete else float(item.change.value or 0), item.threshold or ""])
    return values


def build_failed_excel(failures: list[tuple[PricePlan, str, str]]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Не встановлено"
    headers = ["Статус", "Помилка", "Внутрішній артикул Хорошоп", "Артикул для відображення", "Рядок", "Артикул користувача", "Поточна ціна (грн)", "Поточна ціна (% від РРЦ)", "РРЦ (грн)", "Поточну ціну в РРЦ (Так)", "РРЦ в поточну ціну (Так)"]
    for tier in range(1, 6):
        headers.extend([f"Опт {tier} (грн)", f"Опт {tier} від (шт.)"])
    sheet.append(headers)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=len(headers)).coordinate}"
    for plan, status, message in failures:
        row = plan.rows[0]
        sheet.append([status, message, plan.article, plan.display_article] + price_row_excel_values(row))
    widths = {
        "A": 14, "B": 56, "C": 26, "D": 26, "E": 10, "F": 28, "G": 18, "H": 24, "I": 14, "J": 26, "K": 26,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for index in range(12, len(headers) + 1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = 14
    _style_header(sheet)
    for row in sheet.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def _extract_wholesale_pairs(wholesale_data: Any) -> list[tuple[float, int]]:
    pairs: list[tuple[float, int]] = []
    if isinstance(wholesale_data, (list, tuple)):
        for item in wholesale_data:
            if isinstance(item, dict):
                threshold = item.get("minimal_threshold") or item.get("threshold")
                price = item.get("price")
                if threshold is not None and price is not None:
                    try:
                        pairs.append((float(price), int(threshold)))
                    except (ValueError, TypeError):
                        pass
    pairs.sort(key=lambda pair: pair[1])
    return pairs


def build_report_excel(plans: list[PricePlan]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Детальний звіт цін"
    headers = [
        "Рядок",
        "Артикул у файлі",
        "Артикул сайту",
        "Внутрішній артикул Хорошоп",
        "Стара поточна ціна (грн)",
        "Нова поточна ціна (грн)",
        "Стара РРЦ (грн)",
        "Нова РРЦ (грн)",
    ]
    for tier in range(1, 6):
        headers.extend([
            f"Старий Опт {tier} (грн)",
            f"Старий Опт {tier} від (шт.)",
            f"Новий Опт {tier} (грн)",
            f"Новий Опт {tier} від (шт.)",
        ])
    headers.extend(["Зміни", "Примітки"])
    sheet.append(headers)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=len(headers)).coordinate}"

    for plan in plans:
        row = plan.rows[0]
        product = plan.catalog_product
        user_article = row.display_article
        site_article = (product.display_article if product else "") or plan.display_article
        internal_article = plan.article
        row_number = row.row_number

        old_current = float(product.price) if product and product.price is not None else ""
        new_current = float(plan.payload["price"]) if "price" in plan.payload else old_current

        old_rrc = float(product.price_old) if product and product.price_old is not None else ""
        if "price_old" in plan.payload:
            new_rrc = float(plan.payload["price_old"]) if plan.payload["price_old"] else "Видалено"
        else:
            new_rrc = old_rrc

        old_wholesale_pairs = _extract_wholesale_pairs(product.wholesale_prices) if product else []
        if "wholesale_prices" in plan.payload:
            new_wholesale_pairs = _extract_wholesale_pairs(plan.payload["wholesale_prices"])
        else:
            new_wholesale_pairs = old_wholesale_pairs

        wholesale_values: list[Any] = []
        for tier_idx in range(5):
            old_price = old_wholesale_pairs[tier_idx][0] if tier_idx < len(old_wholesale_pairs) else ""
            old_threshold = old_wholesale_pairs[tier_idx][1] if tier_idx < len(old_wholesale_pairs) else ""
            new_price = new_wholesale_pairs[tier_idx][0] if tier_idx < len(new_wholesale_pairs) else ""
            new_threshold = new_wholesale_pairs[tier_idx][1] if tier_idx < len(new_wholesale_pairs) else ""
            wholesale_values.extend([old_price, old_threshold, new_price, new_threshold])

        changes: list[str] = []
        if "price" in plan.payload:
            if old_current != "" and old_current != new_current:
                changes.append(f"Поточна ціна: {old_current} -> {new_current}")
            else:
                changes.append(f"Поточна ціна: {new_current}")
        if "price_old" in plan.payload:
            if old_rrc != "" and old_rrc != new_rrc:
                changes.append(f"РРЦ: {old_rrc} -> {new_rrc}")
            else:
                changes.append(f"РРЦ: {new_rrc}")
        if "wholesale_prices" in plan.payload:
            changes.append("Опт оновлено")

        sheet.append([
            row_number,
            user_article,
            site_article,
            internal_article,
            old_current,
            new_current,
            old_rrc,
            new_rrc,
        ] + wholesale_values + [
            "; ".join(changes),
            "; ".join(plan.warnings),
        ])

    widths: dict[str, int] = {
        "A": 10, "B": 24, "C": 26, "D": 26, "E": 20, "F": 20, "G": 16, "H": 16,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for index in range(9, len(headers) - 1):
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = 16
    sheet.column_dimensions[sheet.cell(row=1, column=len(headers) - 1).column_letter].width = 35
    sheet.column_dimensions[sheet.cell(row=1, column=len(headers)).column_letter].width = 45

    _style_header(sheet)
    for row_cells in sheet.iter_rows(min_row=2):
        row_cells[-1].alignment = Alignment(wrap_text=True, vertical="top")
        row_cells[-2].alignment = Alignment(wrap_text=True, vertical="top")

    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def _style_header(worksheet: Any) -> None:
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="6D3B8F")
        cell.alignment = Alignment(horizontal="center")


def build_excel_template() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Ціни"
    headers = ["Артикул", "Поточна ціна (грн)", "Поточна ціна (% від РРЦ)", "РРЦ (грн)", "Поточну ціну в РРЦ (Так)", "РРЦ в поточну ціну (Так)"]
    for tier in range(1, 6):
        headers.extend([f"Опт {tier} (грн)", f"Опт {tier} від (шт.)"])
    sheet.append(headers)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = "A1:P1"
    for column, width in {"A": 28, "B": 16, "C": 24, "D": 16, "E": 28, "F": 28, "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14, "M": 14, "N": 14, "O": 14, "P": 14}.items():
        sheet.column_dimensions[column].width = width
    _style_header(sheet)
    examples = [("01063", "", 75, "", "Так", "", "", "", "", "", "", "", 650, 20, "", "")]
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
        "Один рядок = один товар. Кожен стовпець керує окремою ціною цього товару.",
        "Артикул можна вказати як article_for_display зі сайту або внутрішній article Хорошоп. Пошук: точний, потім без урахування регістру.",
        "Порожня клітинка означає «не змінювати». Видалення виконується тільки словом «Видалити» у клітинці РРЦ або потрібного Опт 1-5.",
        "«Поточна ціна, % від РРЦ»: вкажіть число від 1 до 100, у тому числі дробове. Поточна ціна = РРЦ × відсоток / 100 з округленням до цілої гривні вниз.",
        "«Поточну ціну в РРЦ (Так)»: коли вказано нову поточну ціну, попередня поточна ціна буде записана в РРЦ. Не заповнюйте при цьому стовпець РРЦ.",
        "«РРЦ в поточну ціну (Так)» копіює наявне РРЦ у поточну ціну. Щоб одночасно прибрати РРЦ, напишіть «Видалити» у стовпці РРЦ.",
        "Опт 1-3 розраховується від РРЦ/старої ціни за правилами 2 шт. -10%, 5 шт. -15%, 10 шт. -25%. Правила можна змінити у config.json.",
        "Акційна ціна не зменшує опт. Оптові ціни, які не нижчі за поточну ціну, автоматично виключаються. Інші наявні рівні опта зберігаються.",
        "Для будь-якого Опт 1-5 можна вказати «Кількість від». Якщо для Опт 1-3 вона порожня, поріг береться з config.json. Для Опт 4-5 кількість обов'язкова.",
        "Стовпці можна переставляти. Якщо необов'язковий стовпець видалено, він не буде змінюватися. Стовпець «Артикул» обов'язковий.",
    ]
    for index, note in enumerate(notes, start=3):
        guide.cell(index, 1, note).alignment = Alignment(wrap_text=True, vertical="top")
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def extract_clusters(mappings: ArticleMappings) -> list[list[str]]:
    visited: set[str] = set()
    clusters: list[list[str]] = []
    for source in sorted(mappings.entries.keys()):
        if source in visited:
            continue
        cluster: list[str] = []
        queue = [source]
        visited.add(source)
        while queue:
            current = queue.pop(0)
            cluster.append(current)
            for target in mappings.entries.get(current, ()):
                if target not in visited:
                    visited.add(target)
                    queue.append(target)
        if len(cluster) >= 2:
            clusters.append(sorted(cluster))
    clusters.sort(key=lambda item: item[0])
    return clusters


def build_mappings_excel(mappings: ArticleMappings) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Таблиця сопоставлень"

    clusters = extract_clusters(mappings)
    max_cols = max((len(item) for item in clusters), default=2)
    max_cols = max(max_cols, 2)

    headers = [f"Артикул {index}" for index in range(1, max_cols + 1)]
    sheet.append(headers)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=len(headers)).coordinate}"
    _style_header(sheet)

    for row_index, cluster in enumerate(clusters, start=2):
        row_values = list(cluster) + [""] * (max_cols - len(cluster))
        sheet.append(row_values)
        for column_index in range(1, max_cols + 1):
            sheet.cell(row=row_index, column=column_index).number_format = "@"

    for column_index in range(1, max_cols + 1):
        column_letter = sheet.cell(row=1, column=column_index).column_letter
        sheet.column_dimensions[column_letter].width = 30

    register_sheet = workbook.create_sheet("Повний реєстр")
    register_headers = ["Пошуковий артикул", "Кількість відповідностей", "Всі відповідні артикули"]
    register_sheet.append(register_headers)
    register_sheet.freeze_panes = "A2"
    register_sheet.auto_filter.ref = "A1:C1"
    _style_header(register_sheet)

    for row_index, (source, targets) in enumerate(sorted(mappings.entries.items()), start=2):
        register_sheet.append([source, len(targets), ", ".join(targets)])
        register_sheet.cell(row=row_index, column=1).number_format = "@"
        register_sheet.cell(row=row_index, column=3).number_format = "@"

    register_sheet.column_dimensions["A"].width = 28
    register_sheet.column_dimensions["B"].width = 24
    register_sheet.column_dimensions["C"].width = 60

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
        self.by_normalized: dict[str, list[CatalogProduct]] = {}
        for item in products:
            for key in {item.article, item.display_article} - {""}:
                self.by_exact.setdefault(key, []).append(item)
                self.by_folded.setdefault(key.casefold(), []).append(item)
                normalized = normalize_article_code(key)
                if normalized:
                    self.by_normalized.setdefault(normalized.casefold(), []).append(item)

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

    def resolve(self, value: str, use_normalized: bool = False) -> tuple[CatalogProduct | None, str]:
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
        if use_normalized:
            norm_val = normalize_article_code(value)
            zeros = leading_zeros_count(norm_val)
            normalized = norm_val.casefold()
            candidates = self.by_normalized.get(normalized, [])
            valid_candidates = []
            for item in candidates:
                for key in {item.article, item.display_article} - {""}:
                    norm_k = normalize_article_code(key)
                    if norm_k.casefold() == normalized and leading_zeros_count(norm_k) == zeros:
                        valid_candidates.append(item)
                        break
            normalized_matches = list({item.article: item for item in valid_candidates}.values())
            if len(normalized_matches) == 1:
                return normalized_matches[0], ""
            if len(normalized_matches) > 1:
                return None, f"Нормалізований артикул '{norm_val}' для '{value}' не є унікальним."
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


def plan_prices(
    rows: list[PriceRow],
    catalog: CatalogIndex,
    settings: Settings,
    normalize_articles: bool = False,
    mappings: ArticleMappings | None = None,
) -> list[PricePlan]:
    grouped: dict[str, PriceRow] = {}
    grouped_products: dict[str, CatalogProduct] = {}
    errors: list[PricePlan] = []
    mappings = mappings or ArticleMappings.empty()
    for row in rows:
        products: dict[str, CatalogProduct] = {}
        resolution_errors: list[str] = []
        for target_article in mappings.targets_for(row.display_article, normalize_articles):
            product, error = catalog.resolve(target_article, normalize_articles)
            if product is not None:
                products[product.article] = product
            elif error and "не знайдений у каталозі" not in error:
                resolution_errors.append(error)
        if not products:
            error = resolution_errors[0] if resolution_errors else f"Жоден варіант артикула для '{row.display_article}' не знайдений у каталозі."
            errors.append(PricePlan("", row.display_article, {}, (row,), error=error))
            continue
        for product in products.values():
            if product.article in grouped:
                errors.append(PricePlan(product.article, product.display_article or product.article, {}, (row,), error="Артикул повторюється у файлі або базі сопоставлень. Для товару має бути лише одна команда."))
            else:
                grouped[product.article] = row
                grouped_products[product.article] = product
    plans = errors
    for article, row in grouped.items():
        product = grouped_products[article]
        current = product.price
        old = product.price_old
        has_current = row.current.value is not None or row.current_percent_of_old is not None or row.move_old_to_current
        has_old = row.old.specified or row.move_current_to_old
        direct_wholesale = bool(row.wholesale)
        wholesale = wholesale_entries(product.wholesale_prices)
        warnings: list[str] = []
        if row.move_current_to_old:
            old = product.price
        elif row.old.delete:
            old = None
        elif row.old.value is not None:
            old = row.old.value
        if row.move_old_to_current:
            current = product.price_old
        elif row.current_percent_of_old is not None:
            if old is None:
                plans.append(PricePlan(article, product.display_article or article, {}, (row,), error="Немає РРЦ для обчислення поточної ціни за відсотком."))
                continue
            current = (old * row.current_percent_of_old / Decimal("100")).quantize(Decimal("1"), rounding=ROUND_DOWN)
        elif row.current.value is not None:
            current = row.current.value
        elif row.old.value is not None and not has_current:
            if product.price is not None and product.price_old is not None and product.price == product.price_old:
                current = old
                has_current = True
                warnings.append("Поточну ціну встановлено в значення нової РРЦ (до зміни ціни збігалися).")
        if row.move_old_to_current and product.price_old is None:
            plans.append(PricePlan(article, product.display_article or article, {}, (row,), error="Немає РРЦ, яке можна перенести в поточну ціну."))
            continue
        if current is None:
            plans.append(PricePlan(article, product.display_article or article, {}, (row,), error="Немає поточної ціни: спочатку встановіть «Поточна ціна«."))
            continue
        recalculate_wholesale = (has_current or has_old) and settings.recalculate_wholesale_from_rrc and old is not None
        if recalculate_wholesale:
            for rule in settings.wholesale_rules:
                if rule.tier <= 3:
                    wholesale[rule.minimal_threshold] = (old * (Decimal("100") - rule.discount_percent) / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            warnings.append("Опт 1-3 розраховано від РРЦ за правилами за замовчуванням.")
        for wholesale_change in row.wholesale:
            tier = wholesale_change.tier
            rule = rule_for(settings, tier)
            threshold = wholesale_change.threshold or (rule.minimal_threshold if rule else None)
            if threshold is None:
                plans.append(PricePlan(article, product.display_article or article, {}, (row,), error=f"Для Опт {tier} вкажіть «Кількість від»."))
                break
            if rule is not None and threshold != rule.minimal_threshold:
                wholesale.pop(rule.minimal_threshold, None)
            if wholesale_change.change.delete:
                wholesale.pop(threshold, None)
                warnings.append(f"Видалено Опт {tier} від {threshold} шт.")
            else:
                wholesale[threshold] = wholesale_change.change.value  # type: ignore[index]
        else:
            invalid = [threshold for threshold, value in wholesale.items() if value >= current]
            for threshold in invalid:
                del wholesale[threshold]
            if invalid:
                warnings.append("Вимкнено оптові рівні, де ціна не нижча за поточну: " + ", ".join(map(str, sorted(invalid))) + ".")
            payload: dict[str, Any] = {"article": article}
            if has_current: payload["price"] = float(current)
            if has_old: payload["price_old"] = float(old) if old else 0
            if has_current or direct_wholesale or recalculate_wholesale:
                payload["wholesale_prices"] = [{"minimal_threshold": threshold, "price": float(value)} for threshold, value in sorted(wholesale.items())]
            plans.append(PricePlan(article, product.display_article or article, payload, (row,), tuple(warnings), catalog_product=product))
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
