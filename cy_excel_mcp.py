import atexit
import argparse
import json
import os
import re
import time
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from typing import Any
from urllib.parse import quote

import msal
import requests
from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field, model_validator

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - optional dependency
    load_dotenv = None

if load_dotenv is not None:
    load_dotenv()


TENANT_ID = os.getenv("OC_OD_TENANT_ID", "consumers")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.ReadWrite.All"]
CACHE_FILE = os.getenv("OC_OD_CACHE_FILE", "onedrive_token_cache.bin")
FILE_PATH = os.getenv("OC_OD_FILE_PATH", "订单汇总.xlsx")
TABLE_NAME = os.getenv("OC_OD_TABLE_NAME", "表4")
PRODUCT_FILE_PATH = os.getenv("OC_OD_PRODUCT_FILE_PATH", "众一/2026诚亿报表.xlsx")
PRODUCT_SHEET_NAME = os.getenv("OC_OD_PRODUCT_SHEET_NAME", "产品明细")
PRODUCT_NAME_COLUMN = os.getenv("OC_OD_PRODUCT_NAME_COLUMN", "B")
PRODUCT_CATEGORY_COLUMN = os.getenv("OC_OD_PRODUCT_CATEGORY_COLUMN", "C")
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_HOST = os.getenv("CY_EXCEL_MCP_HOST", "127.0.0.1")
DEFAULT_PORT = int(os.getenv("CY_EXCEL_MCP_PORT", "18061"))
LOGS_DIR = os.path.join(os.getcwd(), "logs")
DRAFT_CACHE_FILE = os.path.join(os.getcwd(), os.getenv("CY_EXCEL_MCP_DRAFT_CACHE_FILE", "order_draft_cache.json"))
DRAFT_CACHE_TTL_SECONDS = int(os.getenv("CY_EXCEL_MCP_DRAFT_CACHE_TTL_SECONDS", str(2 * 60 * 60)))
PRODUCT_CACHE_FILE = os.path.join(os.getcwd(), os.getenv("CY_PRODUCT_CACHE_FILE", "product_catalog_cache.json"))
PRODUCT_ALIAS_FILE = os.path.join(os.getcwd(), os.getenv("CY_PRODUCT_ALIAS_FILE", "product_aliases.json"))
DEFAULT_PRODUCT_ALIASES = {
    "98-400": "98400-14oz-400ml",
    "98-400pet": "98400-14oz-400ml",
    "定制98-400杯": "98400-14oz-400ml",
    "98-400ml": "98400-14oz-400ml",
}
EXCEL_HEADERS = [
    "备注",
    "发货厂家",
    "产品供应商",
    "日期",
    "单号",
    "销售员",
    "客户",
    "货品名称",
    "数量",
    "数量单位",
    "销售单价",
    "销售金额",
    "成本单价",
    "成本金额",
    "运费",
    "利润",
    "总货款",
    "已收",
    "未收",
    "收货联系人",
    "收货人电话",
    "收货地址",
]


def _read_text_file(path: str) -> str:
    with open(path, "r", encoding="utf-8") as file:
        return file.read()


def _write_text_file(path: str, content: str) -> None:
    with open(path, "w", encoding="utf-8") as file:
        file.write(content)


def _load_json_file(path: str) -> Any:
    if not os.path.exists(path):
        return None
    try:
        return json.loads(_read_text_file(path))
    except Exception:
        return None


def _write_json_file(path: str, payload: Any) -> None:
    _write_text_file(path, json.dumps(payload, ensure_ascii=False, indent=2))


def _normalize_onedrive_path(path: str | None) -> str:
    normalized = _normalize_value(path)
    if normalized is None:
        return ""
    return str(normalized).replace("\\", "/").strip("/")


def _build_msal_http_client() -> requests.Session:
    session = requests.Session()
    # Avoid inheriting broken system proxy settings during Microsoft login.
    session.trust_env = False
    return session


def _is_network_error(exc: Exception) -> bool:
    error_text = str(exc).lower()
    network_markers = (
        "name or service not known",
        "failed to resolve",
        "temporary failure in name resolution",
        "max retries exceeded",
        "connection aborted",
        "connection refused",
        "proxyerror",
        "ssl",
        "read timed out",
        "connect timeout",
    )
    return any(marker in error_text for marker in network_markers)


def _json_result(**payload: Any) -> str:
    return json.dumps(payload, ensure_ascii=False)


def _append_tool_log(tool_name: str, order_number: Any = None, result_type: Any = None, **extra: Any) -> None:
    os.makedirs(LOGS_DIR, exist_ok=True)
    log_path = os.path.join(LOGS_DIR, f"tool-requests-{datetime.now().strftime('%Y-%m-%d')}.jsonl")
    payload = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "tool_name": tool_name,
        "order_number": _normalize_value(order_number),
        "result_type": _normalize_value(result_type),
    }
    for key, value in extra.items():
        payload[key] = _normalize_value(value)

    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(json.dumps(payload, ensure_ascii=False) + "\n")


def _normalize_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        normalized = value.strip()
        return normalized or None
    return value


def _to_string(value: Any) -> str:
    normalized = _normalize_value(value)
    return "" if normalized is None else str(normalized)


def _normalize_match_text(value: Any) -> str:
    text = _to_string(value).strip()
    if not text:
        return ""
    return re.sub(r"\s+", "", text).casefold()


def _load_recent_draft_records() -> list[dict[str, Any]]:
    payload = _load_json_file(DRAFT_CACHE_FILE)
    return payload if isinstance(payload, list) else []


def _save_recent_draft_records(records: list[dict[str, Any]]) -> None:
    os.makedirs(os.path.dirname(DRAFT_CACHE_FILE) or ".", exist_ok=True)
    _write_json_file(DRAFT_CACHE_FILE, records[-50:])


def _normalize_row_indexes(value: Any) -> list[int]:
    if isinstance(value, list):
        result: list[int] = []
        for item in value:
            try:
                result.append(int(item))
            except (TypeError, ValueError):
                continue
        return result
    return []


def _store_recent_draft(
    order: "ExcelOrder",
    sender_name: str | None,
    row_indexes: list[int] | None = None,
    historical_row_indexes: list[int] | None = None,
    pending_replace: bool = False,
    matched_by: str | None = None,
) -> None:
    records = _load_recent_draft_records()
    normalized_sender = _normalize_match_text(sender_name or order.销售员)
    normalized_order_number = _normalize_match_text(order.单号)
    normalized_customer = _normalize_match_text(order.客户)
    normalized_alias = _normalize_match_text(order.匹配客户别名)

    filtered_records: list[dict[str, Any]] = []
    cutoff = time.time() - DRAFT_CACHE_TTL_SECONDS
    for record in records:
        try:
            created_at = float(record.get("created_at") or 0)
        except (TypeError, ValueError):
            created_at = 0
        if created_at < cutoff:
            continue
        same_sender = _normalize_match_text(record.get("sender_name")) == normalized_sender
        same_order = normalized_order_number and _normalize_match_text(record.get("order_number")) == normalized_order_number
        same_customer = normalized_customer and _normalize_match_text(record.get("customer")) == normalized_customer
        same_alias = normalized_alias and _normalize_match_text(record.get("customer_alias")) == normalized_alias
        if same_sender and (same_order or same_customer or same_alias):
            continue
        filtered_records.append(record)

    filtered_records.append(
        {
            "created_at": time.time(),
            "sender_name": sender_name or order.销售员,
            "order_number": order.单号,
            "customer": order.客户,
            "customer_alias": order.匹配客户别名,
            "row_indexes": row_indexes or [],
            "historical_row_indexes": historical_row_indexes or [],
            "pending_replace": pending_replace,
            "matched_by": matched_by,
            "message_intent": _to_string(order.extra_fields.get("消息意图")),
            "order": order.model_dump(exclude_none=True),
        }
    )
    _save_recent_draft_records(filtered_records)

def _find_recent_draft(
    order: "ExcelOrder",
    sender_name: str | None,
    prefer_pending_replace: bool = False,
) -> dict[str, Any] | None:
    normalized_sender = _normalize_match_text(sender_name or order.销售员)
    normalized_order_number = _normalize_match_text(order.单号)
    normalized_customer = _normalize_match_text(order.客户)
    normalized_alias = _normalize_match_text(order.匹配客户别名)

    best_record: dict[str, Any] | None = None
    best_score = -1
    cutoff = time.time() - DRAFT_CACHE_TTL_SECONDS
    for record in reversed(_load_recent_draft_records()):
        try:
            created_at = float(record.get("created_at") or 0)
        except (TypeError, ValueError):
            continue
        if created_at < cutoff:
            continue
        if normalized_sender and _normalize_match_text(record.get("sender_name")) != normalized_sender:
            continue

        pending_score = 0
        if prefer_pending_replace:
            pending_score = 100 if bool(record.get("pending_replace")) else -100

        score = -1
        if normalized_order_number and _normalize_match_text(record.get("order_number")) == normalized_order_number:
            score = pending_score + 3
        elif normalized_customer and _normalize_match_text(record.get("customer")) == normalized_customer:
            score = pending_score + 2
        elif normalized_alias and _normalize_match_text(record.get("customer_alias")) == normalized_alias:
            score = pending_score + 1

        if score > best_score:
            best_score = score
            best_record = record

    return best_record if best_score >= 0 else None


def _to_float(value: Any) -> float | None:
    normalized = _normalize_value(value)
    if normalized is None:
        return None
    if isinstance(normalized, (int, float)):
        return float(normalized)

    cleaned = str(normalized).replace("元", "").replace(",", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if not match:
        return None
    return float(match.group())


def _format_money(value: float | None) -> str | None:
    if value is None:
        return None
    if value.is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _format_unit_price(value: Any) -> str | None:
    normalized = _normalize_value(value)
    if normalized is None:
        return None

    if isinstance(normalized, (int, float)):
        raw_number = str(normalized)
    else:
        cleaned = str(normalized).replace("元", "").replace(",", "").strip()
        match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
        if not match:
            return None
        raw_number = match.group()

    try:
        decimal_value = Decimal(raw_number)
    except (InvalidOperation, ValueError):
        return None

    formatted = format(decimal_value.normalize(), "f")
    if "." in formatted:
        formatted = formatted.rstrip("0").rstrip(".")
    return _normalize_value(formatted)


def _extract_first(pattern: str, text: str, flags: int = 0) -> str | None:
    match = re.search(pattern, text, flags)
    if not match:
        return None
    return match.group(1).strip()


PHONE_PATTERN = re.compile(r"(?<!\d)(1\d{10})(?!\d)")


def _extract_phone(text: str) -> str | None:
    match = PHONE_PATTERN.search(text)
    return match.group(1) if match else None


def _same_text(left: str | None, right: str | None) -> bool:
    return _to_string(left) == _to_string(right)


def _normalize_entity_name(value: str | None) -> str | None:
    normalized = _normalize_value(value)
    if normalized is None:
        return None
    cleaned = unicodedata.normalize("NFKC", str(normalized)).strip()
    cleaned = re.sub(r"^[\"'“”‘’\s]+|[\"'“”‘’\s]+$", "", cleaned)
    cleaned = cleaned.strip(" ：:，,；;。.\t\r\n")
    return _normalize_value(cleaned)


def _extract_order_number_and_customer_alias(text: str) -> tuple[str | None, str | None]:
    raw_order_number = _extract_first(r"单号[:：]\s*([^\r\n]+)", text)
    if raw_order_number is None:
        return None, None

    normalized_order_number = unicodedata.normalize("NFKC", raw_order_number).strip(" ：:，,；;。.\t\r\n")
    alias = None
    alias_match = re.search(r"[（(]([^()（）]+)[)）]\s*[，,；;。.\s]*$", normalized_order_number)
    if alias_match:
        alias = _normalize_entity_name(alias_match.group(1))
        normalized_order_number = re.sub(
            r"\s*[（(][^()（）]+[)）]\s*[，,；;。.\s]*$",
            "",
            normalized_order_number,
        ).strip(" ：:，,；;。.\t\r\n")

    return _normalize_entity_name(normalized_order_number), alias


def _normalize_ocr_line(line: str) -> str:
    return re.sub(r"\s+", " ", line.strip())


def _is_noise_line(line: str) -> bool:
    cleaned = _normalize_ocr_line(line)
    if not cleaned:
        return True

    lower_cleaned = cleaned.lower()
    exact_noise = {
        "客户",
        "数量",
        "单价",
        "金额",
        "合计",
        "共计",
        "已收",
    }
    if cleaned in exact_noise:
        return True

    keyword_prefixes = (
        "单号",
        "收件人",
        "收货人",
        "收货联系人",
        "手机",
        "手机号码",
        "电话号码",
        "电话",
        "收货地址",
        "地址",
        "所在地区",
        "所在地址",
        "地区",
        "全款",
        "总货款",
        "合计",
        "共计",
        "已收",
        "客户:",
        "客户：",
        "备注",
    )
    if cleaned.startswith(keyword_prefixes):
        return True

    if PHONE_PATTERN.search(cleaned):
        return True

    if lower_cleaned.startswith("item ") or lower_cleaned.startswith("product "):
        return False

    if re.fullmatch(r"共?\s*\d+(?:\.\d+)?\s*元", cleaned):
        return True

    if "收款" in cleaned and len(cleaned) <= 20:
        return True

    return False


def _normalize_replace_target(text: str | None) -> str | None:
    normalized = _normalize_value(text)
    if normalized is None:
        return None

    cleaned = str(normalized).strip()
    replace_suffixes = (
        "以这个为准",
        "这个为准",
        "以这个为主",
        "这个为主",
        "以上面这个为准",
        "以下图为准",
        "按这个为准",
        "前面的作废",
        "之前那个作废",
        "修改订单",
    )
    for suffix in replace_suffixes:
        if cleaned.endswith(suffix):
            cleaned = cleaned[: -len(suffix)].strip(" ：:，,；;。.")
            break
    return _normalize_entity_name(cleaned)


def _extract_customer_name(text: str) -> str | None:
    customer = _extract_first(r"客户[:：]\s*([^\r\n]+)", text)
    if customer:
        return _normalize_replace_target(customer)

    lines = [_normalize_ocr_line(line) for line in text.splitlines() if line.strip()]
    for line in lines:
        if _is_noise_line(line):
            continue
        if _looks_like_address(line):
            continue
        if re.search(r"\d", line):
            continue
        return _normalize_replace_target(line)
    return None


def _clean_contact_name(value: str | None) -> str | None:
    normalized = _normalize_entity_name(value)
    if normalized is None:
        return None

    cleaned = PHONE_PATTERN.sub("", normalized)
    cleaned = re.sub(
        r"^(?:收件人|收货人|收货联系人|联系人|姓名|手机号码|手机号|电话号码|电话)[:：]?\s*",
        "",
        cleaned,
    )
    cleaned = cleaned.strip(" ：:，,；;。.\t\r\n")
    if not cleaned:
        return None
    if "收款" in cleaned:
        return None
    if re.search(r"\d", cleaned):
        return None
    if _looks_like_address(cleaned):
        return None
    if len(cleaned) > 30:
        return None
    return _normalize_entity_name(cleaned)


def _extract_contact_name(text: str) -> str | None:
    explicit = (
        _extract_first(r"收件人[:：]\s*([^\r\n]+)", text)
        or _extract_first(r"收货人[:：]\s*([^\r\n]+)", text)
        or _extract_first(r"收货联系人[:：]\s*([^\r\n]+)", text)
    )
    if explicit:
        contact = _clean_contact_name(explicit)
        if contact:
            return contact

    lines = [_normalize_ocr_line(line) for line in text.splitlines() if line.strip()]
    for line in lines:
        if not PHONE_PATTERN.search(line):
            continue

        after_phone = re.search(r"(?<!\d)1\d{10}(?!\d)[，,；;\s]+([^\r\n]+)", line)
        if after_phone:
            contact = _clean_contact_name(after_phone.group(1))
            if contact:
                return contact

        before_phone = re.search(r"([^\r\n，,；;:：]{1,30})[，,；;\s]+(?<!\d)1\d{10}(?!\d)", line)
        if before_phone:
            contact = _clean_contact_name(before_phone.group(1))
            if contact:
                return contact

    for index, line in enumerate(lines):
        if not PHONE_PATTERN.search(line):
            continue
        for neighbor_index in (index - 1, index + 1):
            if neighbor_index < 0 or neighbor_index >= len(lines):
                continue
            contact = _clean_contact_name(lines[neighbor_index])
            if contact:
                return contact
    return None


def _extract_salesperson_name(text: str) -> str | None:
    explicit = (
        _extract_first(r"销售员[:：]\s*(.+)", text)
        or _extract_first(r"业务员[:：]\s*(.+)", text)
    )
    return _normalize_value(explicit)


def _looks_like_address(line: str) -> bool:
    normalized = _normalize_ocr_line(line)
    if not normalized:
        return False
    if re.search(r"1\d{10}", normalized):
        return False
    if "收款" in normalized:
        return False
    if _parse_item_from_structured_line(normalized) is not None:
        return False
    if _parse_item_from_table_line(normalized) is not None:
        return False
    address_keywords = ("省", "市", "区", "县", "镇", "乡", "村", "街道", "大道", "路", "街", "号", "栋", "幢", "楼", "单元", "室")
    keyword_hits = sum(1 for keyword in address_keywords if keyword in normalized)
    return keyword_hits >= 2


def _extract_address(text: str) -> str | None:
    explicit = _extract_first(r"(?:收货地址|地址|所在地区|地区|所在地址)[:：]\s*(.+)", text)
    if explicit:
        return explicit

    candidate_lines = [_normalize_ocr_line(line) for line in text.splitlines() if line.strip()]
    address_candidates = [line for line in candidate_lines if _looks_like_address(line)]
    if not address_candidates:
        return None
    address_candidates.sort(key=len, reverse=True)
    return _normalize_value(address_candidates[0])


def _split_quantity_and_unit(raw: str | None) -> tuple[str | None, str | None]:
    normalized = _normalize_value(raw)
    if normalized is None:
        return None, None
    match = re.match(r"(?P<qty>\d+(?:\.\d+)?)(?P<unit>.*)", str(normalized).strip())
    if not match:
        return _normalize_value(normalized), None
    qty = _normalize_value(match.group("qty"))
    unit = _normalize_quantity_unit(match.group("unit"))
    return qty, unit


def _normalize_quantity_unit(raw: str | None) -> str | None:
    normalized = _normalize_value(raw)
    if normalized is None:
        return None
    unit_text = str(normalized).strip()
    if "(" in unit_text:
        unit_text = unit_text.split("(", 1)[0].strip()
    if "（" in unit_text:
        unit_text = unit_text.split("（", 1)[0].strip()
    return _normalize_value(unit_text)


def _is_plate_fee_name(value: Any) -> bool:
    return "版费" in _to_string(value)


class OrderItem(BaseModel):
    货品名称: str
    数量: str | None = None
    数量单位: str | None = None
    销售单价: str | None = None
    销售金额: str | None = None

    def normalized(self) -> "OrderItem":
        return OrderItem(
            货品名称=_to_string(self.货品名称),
            数量=_normalize_value(self.数量),
            数量单位=_normalize_quantity_unit(self.数量单位),
            销售单价=_format_unit_price(self.销售单价),
            销售金额=_format_money(_to_float(self.销售金额)),
        )

    def signature(self) -> tuple[str, str, str, str, str]:
        normalized = self.normalized()
        return (
            _to_string(normalized.货品名称),
            _to_string(normalized.数量),
            _to_string(normalized.数量单位),
            _to_string(normalized.销售单价),
            _to_string(normalized.销售金额),
        )


def _parse_plate_fee_item_from_line(line: str) -> OrderItem | None:
    normalized = _normalize_ocr_line(line)
    if "版费" not in normalized:
        return None

    fee_text = normalized[normalized.find("版费") :]
    money_matches = re.findall(r"(?:￥|¥)?\s*(\d+(?:\.\d+)?)\s*元", fee_text)
    number_matches = re.findall(r"\d+(?:\.\d+)?", fee_text)
    amount = money_matches[-1] if money_matches else (number_matches[-1] if number_matches else None)
    if amount is None:
        return None

    return OrderItem(
        货品名称="版费",
        数量="1",
        数量单位="项",
        销售单价=amount,
        销售金额=amount,
    ).normalized()


def _parse_item_from_structured_line(line: str) -> OrderItem | None:
    normalized = _normalize_ocr_line(line)
    if not normalized:
        return None

    plate_fee_item = _parse_plate_fee_item_from_line(normalized)
    if plate_fee_item is not None:
        return plate_fee_item

    stripped = re.sub(r"^(商品|货品|产品|item|product)\s*\d*\s*[:：]\s*", "", normalized, flags=re.IGNORECASE)
    if stripped != normalized and ("数量" in stripped or "金额" in stripped):
        name = _extract_first(r"^(.*?)\s*(?:[|｜]|数量[:：])", stripped)
        quantity_raw = _extract_first(r"数量[:：]\s*([^|｜]+)", stripped)
        unit_price = _extract_first(r"单价[:：]\s*([^|｜]+)", stripped)
        amount = _extract_first(r"金额[:：]\s*([^|｜]+)", stripped)
        qty, unit = _split_quantity_and_unit(quantity_raw)
        if name:
            return OrderItem(
                货品名称=name.strip(),
                数量=qty,
                数量单位=unit,
                销售单价=unit_price,
                销售金额=amount,
            ).normalized()
    return None


def _parse_item_from_table_line(line: str) -> OrderItem | None:
    normalized = _normalize_ocr_line(line)
    if _is_noise_line(normalized):
        return None

    plate_fee_item = _parse_plate_fee_item_from_line(normalized)
    if plate_fee_item is not None:
        return plate_fee_item

    match = re.match(
        r"^(?P<name>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s*(?P<unit>[^\d\s]+(?:\([^)]*\))?)?\s+"
        r"(?P<unit_price>\d+(?:\.\d+)?)\s*元?\s+(?P<amount>\d+(?:\.\d+)?)\s*元?$",
        normalized,
    )
    if match:
        return OrderItem(
            货品名称=match.group("name").strip(),
            数量=match.group("qty").strip(),
            数量单位=_normalize_quantity_unit(match.group("unit")),
            销售单价=match.group("unit_price").strip(),
            销售金额=match.group("amount").strip(),
        ).normalized()

    match_amount_only = re.match(
        r"^(?P<name>.+?)(?P<qty>\d+(?:\.\d+)?)\s*(?P<unit>[^\d\s元]+(?:\([^)]*\))?)?\s*(?P<amount>\d+(?:\.\d+)?)\s*元?$",
        normalized,
    )
    if match_amount_only:
        return OrderItem(
            货品名称=match_amount_only.group("name").strip(),
            数量=match_amount_only.group("qty").strip(),
            数量单位=_normalize_quantity_unit(match_amount_only.group("unit")),
            销售金额=match_amount_only.group("amount").strip(),
        ).normalized()
    return None


def _extract_order_items(text: str) -> list[OrderItem]:
    items: list[OrderItem] = []
    seen_signatures: set[tuple[str, str, str, str, str]] = set()

    for raw_line in text.splitlines():
        normalized = _normalize_ocr_line(raw_line)
        if not normalized:
            continue

        item = _parse_item_from_structured_line(normalized)
        if item is None:
            item = _parse_item_from_table_line(normalized)
        if item is None:
            continue

        signature = item.signature()
        if signature in seen_signatures:
            continue
        seen_signatures.add(signature)
        items.append(item)

    return items


def _aggregate_items_for_excel(items: list[OrderItem]) -> dict[str, Any]:
    if not items:
        return {}

    normalized_items = [item.normalized() for item in items]
    total_amount = sum(_to_float(item.销售金额) or 0.0 for item in normalized_items)
    has_any_amount = any(_to_float(item.销售金额) is not None for item in normalized_items)

    return {
        "货品名称": "；".join(_to_string(item.货品名称) for item in normalized_items if _to_string(item.货品名称)),
        "数量": "；".join(_to_string(item.数量) for item in normalized_items if _to_string(item.数量)),
        "数量单位": "；".join(_to_string(item.数量单位) for item in normalized_items if _to_string(item.数量单位)),
        "销售单价": "；".join(_to_string(item.销售单价) for item in normalized_items if _to_string(item.销售单价)),
        "销售金额": _format_money(total_amount) if has_any_amount else None,
        "items": [item.model_dump(exclude_none=True) for item in normalized_items],
    }


def _build_excel_row_dicts(order: "ExcelOrder") -> list[dict[str, Any]]:
    finalized_order = order.model_copy(deep=True)
    finalized_order.finalize()

    if not finalized_order.items:
        return [finalized_order.to_excel_dict()]

    total_received = _normalize_value(finalized_order.已收)
    total_unpaid = _normalize_value(finalized_order.未收)
    row_dicts: list[dict[str, Any]] = []

    for index, item in enumerate(finalized_order.items, start=1):
        normalized_item = item.normalized()
        row_order = finalized_order.model_copy(deep=True)
        row_order.items = []
        row_order.货品名称 = _normalize_value(normalized_item.货品名称)
        row_order.数量 = _normalize_value(normalized_item.数量)
        row_order.数量单位 = _normalize_value(normalized_item.数量单位)
        row_order.销售单价 = _normalize_value(normalized_item.销售单价)
        line_amount = _normalize_value(normalized_item.销售金额)
        row_order.销售金额 = line_amount
        row_order.总货款 = line_amount
        # 当前 Graph Table API 不支持稳定的表内合并单元格，这里只在首行保留整单已收/未收。
        row_order.extra_fields["商品序号"] = str(index)
        if _normalize_value(normalized_item.销售金额) is not None:
            row_order.extra_fields["明细金额"] = _normalize_value(normalized_item.销售金额)
        row_dict = row_order.to_excel_dict()
        if index == 1:
            row_dict["已收"] = total_received
            row_dict["未收"] = total_unpaid
            if total_received is not None:
                row_dict["整单已收"] = total_received
            if total_unpaid is not None:
                row_dict["整单未收"] = total_unpaid
        else:
            row_dict["已收"] = ""
            row_dict["未收"] = ""
        row_dicts.append(row_dict)

    return row_dicts


def _has_item_details(order: "ExcelOrder") -> bool:
    if order.items:
        return True
    return any(
        _normalize_value(getattr(order, field_name)) is not None
        for field_name in ("货品名称", "数量", "数量单位", "销售单价")
    )


def _detect_message_intent(raw_message: str, has_item_details: bool) -> str:
    normalized = re.sub(r"\s+", "", raw_message)
    replace_keywords = (
        "以这个为准",
        "这个为准",
        "以这个为主",
        "这个为主",
        "以上面这个为准",
        "以下图为准",
        "按这个为准",
        "前面的作废",
        "之前那个作废",
        "修改订单",
    )
    if any(keyword in normalized for keyword in replace_keywords):
        return "replace_order"
    if has_item_details:
        return "revise_items"
    return "supplement"


# Long-term order-change mode / 长周期改单模式
# - 短周期（图片先发、文字后补）仍优先复用最近草稿缓存。
# - 长周期改单时，不能只依赖最近草稿；当业务员引用旧图片/旧文本并补一句
#   “客户名以这个为准 / 这个为准客户名”时，应优先扫描 Excel 历史订单块。
# - 主匹配锚点应是：销售员 + 客户 + 匹配客户别名；单号存在时单号优先。
# - 一旦明确是 replace_order，就应把命中的整块历史订单视为可替换块，
#   而不是继续要求新旧明细行数一致。
# - 最近草稿缓存只作为短期加速器；若缓存已过期、缺失、或与历史块冲突，
#   应回退到 Excel 历史匹配并记录替换审计信息。


def _normalize_date(date_text: str | None, message_time: str | None) -> str | None:
    source = _normalize_value(date_text)
    if source is None and message_time:
        source = _normalize_value(message_time)
    if source is None:
        return None

    candidates = [
        ("%Y-%m-%d", source),
        ("%Y/%m/%d", source),
        ("%Y.%m.%d", source),
        ("%y.%m.%d", source),
        ("%Y-%m-%d %H:%M:%S", source),
        ("%Y/%m/%d %H:%M:%S", source),
    ]
    for fmt, raw in candidates:
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    match = re.search(r"(\d{2,4})[./-](\d{1,2})[./-](\d{1,2})", source)
    if not match:
        compact_match = re.search(r"(?<!\d)(\d{2})(\d{2})(\d{2})(?:-\d+)?(?!\d)", source)
        if compact_match:
            year = 2000 + int(compact_match.group(1))
            month = int(compact_match.group(2))
            day = int(compact_match.group(3))
            return f"{year:04d}-{month:02d}-{day:02d}"
        return None

    year = int(match.group(1))
    if year < 100:
        year += 2000
    month = int(match.group(2))
    day = int(match.group(3))
    return f"{year:04d}-{month:02d}-{day:02d}"


def _merge_prefer_new(old_value: Any, new_value: Any) -> Any:
    normalized_new = _normalize_value(new_value)
    if normalized_new is not None:
        return normalized_new
    return _normalize_value(old_value)


def _build_token_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(_read_text_file(CACHE_FILE))
    return cache


def _load_cache_payload() -> dict[str, Any] | None:
    if not os.path.exists(CACHE_FILE):
        return None
    try:
        return json.loads(_read_text_file(CACHE_FILE))
    except Exception:
        return None


def _save_cache_payload(payload: dict[str, Any]) -> None:
    try:
        _write_text_file(CACHE_FILE, json.dumps(payload))
    except Exception:
        pass


def _get_valid_cached_access_token(cache_payload: dict[str, Any] | None) -> str | None:
    if not cache_payload:
        return None

    now = int(time.time())
    minimum_expiry = now + 300
    for token_entry in cache_payload.get("AccessToken", {}).values():
        target = str(token_entry.get("target") or "")
        scopes = set(target.split())
        if not set(SCOPES).issubset(scopes):
            continue
        try:
            expires_on = int(token_entry.get("expires_on") or 0)
        except (TypeError, ValueError):
            expires_on = 0
        if expires_on >= minimum_expiry and token_entry.get("secret"):
            return str(token_entry["secret"])
    return None


def _get_cached_refresh_token(cache_payload: dict[str, Any] | None) -> str | None:
    if not cache_payload:
        return None
    for token_entry in cache_payload.get("RefreshToken", {}).values():
        secret = token_entry.get("secret")
        if secret:
            return str(secret)
    return None


def _update_cache_payload_with_token_response(
    cache_payload: dict[str, Any] | None,
    token_response: dict[str, Any],
) -> None:
    if not cache_payload:
        return

    access_tokens = cache_payload.get("AccessToken") or {}
    refresh_tokens = cache_payload.get("RefreshToken") or {}
    now = int(time.time())
    expires_in = int(token_response.get("expires_in") or 3600)
    ext_expires_in = int(token_response.get("ext_expires_in") or expires_in)
    scope_text = token_response.get("scope") or " ".join(SCOPES + ["openid", "profile"])

    for token_entry in access_tokens.values():
        token_entry["secret"] = token_response.get("access_token")
        token_entry["cached_at"] = str(now)
        token_entry["expires_on"] = str(now + expires_in)
        token_entry["extended_expires_on"] = str(now + ext_expires_in)
        token_entry["target"] = scope_text
        break

    new_refresh_token = token_response.get("refresh_token")
    if new_refresh_token:
        for token_entry in refresh_tokens.values():
            token_entry["secret"] = new_refresh_token
            break

    _save_cache_payload(cache_payload)


def _refresh_access_token_from_cache(cache_payload: dict[str, Any] | None) -> str | None:
    refresh_token = _get_cached_refresh_token(cache_payload)
    client_id = os.getenv("OC_OD_CLIENT_ID")
    if not refresh_token or not client_id:
        return None

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    session = _build_msal_http_client()
    response = session.post(
        token_url,
        data={
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": " ".join(SCOPES + ["openid", "profile", "offline_access"]),
        },
        timeout=30,
    )
    if response.status_code != 200:
        return None

    token_response = response.json()
    access_token = token_response.get("access_token")
    if not access_token:
        return None

    _update_cache_payload_with_token_response(cache_payload, token_response)
    return str(access_token)


def _register_cache_persistence(cache: msal.SerializableTokenCache) -> None:
    def _persist_cache() -> None:
        if cache.has_state_changed:
            _write_text_file(CACHE_FILE, cache.serialize())

    atexit.register(_persist_cache)


def _build_public_client_application(
    token_cache: msal.SerializableTokenCache | None = None,
) -> msal.PublicClientApplication | None:
    client_id = os.getenv("OC_OD_CLIENT_ID")
    if not client_id:
        return None

    return msal.PublicClientApplication(
        client_id,
        authority=AUTHORITY,
        token_cache=token_cache,
        http_client=_build_msal_http_client(),
        instance_discovery=False,
    )


def get_token_automatically() -> str | None:
    cache_payload = _load_cache_payload()
    cached_access_token = _get_valid_cached_access_token(cache_payload)
    if cached_access_token:
        return cached_access_token

    try:
        refreshed_access_token = _refresh_access_token_from_cache(cache_payload)
        if refreshed_access_token:
            return refreshed_access_token
    except Exception:
        pass

    cache = _build_token_cache()
    _register_cache_persistence(cache)

    app = _build_public_client_application(token_cache=cache)
    if app is None:
        return None

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "message" not in flow:
        return None

    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    return result.get("access_token")


def _build_base_url() -> str:
    return f"{GRAPH_ROOT}/me/drive/root:/{_normalize_onedrive_path(FILE_PATH)}:/workbook/tables('{TABLE_NAME}')"


def _column_index_to_letter(index: int) -> str:
    result = ""
    current = index + 1
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _column_letters_to_index(letters: str) -> int:
    normalized = _to_string(letters).upper().replace("$", "")
    if not re.fullmatch(r"[A-Z]+", normalized):
        raise ValueError(f"Invalid Excel column letters: {letters}")
    index = 0
    for char in normalized:
        index = index * 26 + (ord(char) - 64)
    return index - 1


def _parse_table_range_address(address: str) -> tuple[str, int, int] | None:
    normalized = address.strip()
    match = re.match(r"^(?P<sheet>.+)!(?P<start_col>\$?[A-Z]+)\$?(?P<start_row>\d+):", normalized)
    if not match:
        return None
    sheet_name = match.group("sheet").strip().strip("'")
    start_col_letters = match.group("start_col").replace("$", "")
    start_row = int(match.group("start_row"))

    start_col_index = _column_letters_to_index(start_col_letters)
    return sheet_name, start_col_index, start_row


def _get_table_layout(
    base_url: str,
    headers: dict[str, str],
    proxies: dict[str, str | None],
) -> tuple[str, int, int, str | None] | None:
    resp_range = requests.get(f"{base_url}/range", headers=headers, proxies=proxies, timeout=30)
    if resp_range.status_code != 200:
        return None
    address = resp_range.json().get("address")
    if not address:
        return None
    parsed = _parse_table_range_address(address)
    if parsed is None:
        return None

    worksheet_id = None
    resp_worksheet = requests.get(f"{base_url}/worksheet", headers=headers, proxies=proxies, timeout=30)
    if resp_worksheet.status_code == 200:
        worksheet_id = resp_worksheet.json().get("id")

    sheet_name, start_col_index, start_row = parsed
    return sheet_name, start_col_index, start_row, worksheet_id


def _format_payment_block(
    base_url: str,
    headers: dict[str, str],
    proxies: dict[str, str | None],
    existing_columns: list[str],
    row_indexes: list[int],
) -> tuple[bool, str | None]:
    layout = _get_table_layout(base_url, headers, proxies)
    if layout is None:
        return False, "无法定位 Excel 表格所在工作表，跳过已收/未收样式设置。"

    sheet_name, start_col_index, header_row, worksheet_id = layout
    data_start_row = header_row + 1
    first_row = data_start_row + min(row_indexes)
    last_row = data_start_row + max(row_indexes)

    def _build_range_bases(column_letter: str) -> list[str]:
        simple_range = f"{column_letter}{first_row}:{column_letter}{last_row}"
        bases: list[str] = []
        if worksheet_id:
            bases.append(
                f"{GRAPH_ROOT}/me/drive/root:/{FILE_PATH}:/workbook/worksheets/"
                f"{quote(worksheet_id, safe='')}/range(address='{simple_range}')"
            )
        bases.append(
            f"{GRAPH_ROOT}/me/drive/root:/{FILE_PATH}:/workbook/worksheets/"
            f"{quote(sheet_name, safe='')}/range(address='{simple_range}')"
        )
        return bases

    for column_name in ("已收", "未收"):
        if column_name not in existing_columns:
            continue
        column_index = start_col_index + existing_columns.index(column_name)
        column_letter = _column_index_to_letter(column_index)
        last_error = None
        styled = False
        for range_base in _build_range_bases(column_letter):
            font_resp = requests.patch(
                f"{range_base}/format/font",
                headers=headers,
                json={"color": "#FF0000", "bold": True},
                proxies=proxies,
                timeout=30,
            )
            if font_resp.status_code != 200:
                last_error = f"{column_name} 列字体着色失败（status={font_resp.status_code}）。"
                continue

            alignment_resp = requests.patch(
                f"{range_base}/format",
                headers=headers,
                json={
                    "horizontalAlignment": "Center",
                    "verticalAlignment": "Center",
                    "wrapText": True,
                },
                proxies=proxies,
                timeout=30,
            )
            if alignment_resp.status_code != 200:
                last_error = f"{column_name} 列对齐设置失败（status={alignment_resp.status_code}）。"
                continue

            fill_resp = requests.patch(
                f"{range_base}/format/fill",
                headers=headers,
                json={"color": "#FFF2CC"},
                proxies=proxies,
                timeout=30,
            )
            if fill_resp.status_code != 200:
                last_error = f"{column_name} 列背景着色失败（status={fill_resp.status_code}）。"
                continue

            border_payload = {
                "style": "Continuous",
                "color": "#C00000",
                "weight": "Thin",
            }
            border_ok = True
            for side in ("EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"):
                border_resp = requests.patch(
                    f"{range_base}/format/borders/{side}",
                    headers=headers,
                    json=border_payload,
                    proxies=proxies,
                    timeout=30,
                )
                if border_resp.status_code != 200:
                    last_error = f"{column_name} 列边框设置失败（status={border_resp.status_code}）。"
                    border_ok = False
                    break
            if not border_ok:
                continue

            styled = True
            break

        if not styled:
            return False, last_error or f"{column_name} 列样式设置失败。"

    return True, None


def _format_unit_price_cells(
    base_url: str,
    headers: dict[str, str],
    proxies: dict[str, str | None],
    existing_columns: list[str],
    row_indexes: list[int],
) -> tuple[bool, str | None]:
    if "销售单价" not in existing_columns or not row_indexes:
        return True, None

    layout = _get_table_layout(base_url, headers, proxies)
    if layout is None:
        return False, "无法定位 Excel 表格所在工作表，跳过销售单价显示格式设置。"

    sheet_name, start_col_index, header_row, worksheet_id = layout
    data_start_row = header_row + 1
    first_row = data_start_row + min(row_indexes)
    last_row = data_start_row + max(row_indexes)
    column_index = start_col_index + existing_columns.index("销售单价")
    column_letter = _column_index_to_letter(column_index)
    simple_range = f"{column_letter}{first_row}:{column_letter}{last_row}"
    row_count = last_row - first_row + 1
    payload = {"numberFormat": [["0.######"] for _ in range(row_count)]}

    range_bases: list[str] = []
    if worksheet_id:
        range_bases.append(
            f"{GRAPH_ROOT}/me/drive/root:/{FILE_PATH}:/workbook/worksheets/"
            f"{quote(worksheet_id, safe='')}/range(address='{simple_range}')"
        )
    range_bases.append(
        f"{GRAPH_ROOT}/me/drive/root:/{FILE_PATH}:/workbook/worksheets/"
        f"{quote(sheet_name, safe='')}/range(address='{simple_range}')"
    )

    last_error = None
    for range_base in range_bases:
        resp_patch = requests.patch(
            range_base,
            headers=headers,
            json=payload,
            proxies=proxies,
            timeout=30,
        )
        if resp_patch.status_code == 200:
            return True, None
        last_error = f"销售单价显示格式设置失败（status={resp_patch.status_code}）。"

    return False, last_error or "销售单价显示格式设置失败。"


def _format_order_rows(
    base_url: str,
    headers: dict[str, str],
    proxies: dict[str, str | None],
    existing_columns: list[str],
    row_indexes: list[int],
) -> tuple[bool, str | None]:
    warnings: list[str] = []
    for formatter in (_format_payment_block, _format_unit_price_cells):
        ok, warning = formatter(
            base_url=base_url,
            headers=headers,
            proxies=proxies,
            existing_columns=existing_columns,
            row_indexes=row_indexes,
        )
        if not ok and warning:
            warnings.append(warning)

    if warnings:
        return False, "；".join(warnings)
    return True, None


def _delete_table_rows(
    base_url: str,
    headers: dict[str, str],
    proxies: dict[str, str | None],
    row_indexes: list[int],
) -> tuple[bool, str | None]:
    for row_index in sorted(row_indexes, reverse=True):
        delete_attempts = (
            ("delete", f"{base_url}/rows/itemAt(index={row_index})", None),
            ("delete", f"{base_url}/rows/{row_index}", None),
            ("post", f"{base_url}/rows/itemAt(index={row_index})/delete", {}),
        )
        last_response: requests.Response | None = None
        for method, url, payload in delete_attempts:
            if method == "delete":
                resp_delete = requests.delete(
                    url,
                    headers=headers,
                    proxies=proxies,
                    timeout=30,
                )
            else:
                resp_delete = requests.post(
                    url,
                    headers=headers,
                    json=payload,
                    proxies=proxies,
                    timeout=30,
                )
            last_response = resp_delete
            if resp_delete.status_code in (200, 202, 204):
                break
        else:
            error_type, error_message = _classify_graph_error(
                last_response.status_code if last_response is not None else 500,
                last_response.text if last_response is not None else "delete row failed",
                operation="update",
            )
            return False, _json_result(
                success=False,
                action="delete_failed",
                error_type=error_type,
                message=error_message,
                row_index=row_index,
                row_indexes=row_indexes,
                status_code=last_response.status_code if last_response is not None else 500,
                detail=last_response.text if last_response is not None else "delete row failed",
            )
    return True, None


def _build_workbook_base_url(file_path: str | None = None) -> str:
    target_file_path = _normalize_onedrive_path(file_path or FILE_PATH)
    return f"{GRAPH_ROOT}/me/drive/root:/{target_file_path}:/workbook"


def _build_drive_item_urls(file_path: str) -> list[str]:
    target_file_path = _normalize_onedrive_path(file_path)
    return [
        f"{GRAPH_ROOT}/me/drive/root:/{target_file_path}",
        f"{GRAPH_ROOT}/me/drive/root:/{target_file_path}:",
    ]


def _graph_headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def _normalize_product_key(value: Any) -> str:
    text = _to_string(value)
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text).casefold()
    text = text.replace("毫升", "ml")
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[·•,，.。;；:：_/\|\\()\[\]{}（）【】<>《》\"'`~!！?？+]", "", text)
    text = text.replace("-", "")
    text = text.replace("订制", "定制")
    return text


def _product_numeric_codes(value: Any) -> set[str]:
    text = unicodedata.normalize("NFKC", _to_string(value))
    numbers = [item.replace(".", "") for item in re.findall(r"\d+(?:\.\d+)?", text)]
    codes = {number for number in numbers if len(number) >= 4}
    if len(numbers) >= 2:
        codes.add(numbers[0].lstrip("0") + numbers[1].lstrip("0"))
    return {code for code in codes if code}


def _category_family(value: Any) -> str:
    category = _normalize_product_key(value)
    if not category:
        return ""
    if "盖" in category:
        if "注塑" in category:
            return "注塑盖"
        if "pet" in category:
            return "PET盖"
        return "盖"
    if "杯" in category or "pet" == category:
        if "纸" in category:
            return "纸杯"
        if "注塑" in category:
            return "注塑杯"
        if "吸塑" in category:
            return "吸塑杯"
        if "pet" in category:
            return "PET杯"
        return "杯"
    if "吸管" in category or "管" == category:
        if "降解" in category:
            return "可降解吸管"
        return "吸管"
    if "袋" in category:
        if "纸" in category:
            return "纸袋"
        if "无纺布" in category:
            return "无纺布袋"
        return "袋"
    if "防漏纸" in category:
        return "防漏纸"
    if "膜" in category:
        return "膜"
    if "费用" in category or "费" in category or "定金" in category:
        return "费用"
    return category


def _infer_product_family(raw_name: Any) -> str:
    text = _normalize_product_key(raw_name)
    if not text:
        return ""
    if "盖" in text:
        if "注塑" in text:
            return "注塑盖"
        if "pet" in text or "直饮" in text or "拱" in text or "平" in text or "外扣" in text:
            return "PET盖"
        return "盖"
    if "吸管" in text or re.search(r"(大管|小管|粗管|细管|单支.*管)", text):
        return "可降解吸管" if "降解" in text else "吸管"
    if "袋" in text:
        if "纸" in text or "牛皮" in text or "白皮" in text:
            return "纸袋"
        if "无纺布" in text or "保温袋" in text:
            return "无纺布袋"
        return "袋"
    if "防漏纸" in text or "锡纸" in text:
        return "防漏纸"
    if "纸杯" in text or "淋膜" in text or "单层" in text or "双层" in text:
        return "纸杯"
    if "注塑" in text or "鸳鸯杯" in text or "方瓶" in text or "膜内贴" in text:
        return "注塑杯"
    if "吸塑" in text:
        return "吸塑杯"
    if "pet" in text:
        return "PET杯"
    if "杯" in text or "ml" in text or "oz" in text or re.search(r"^\d{2,3}[-]?\d{2,4}", text):
        return "杯"
    return ""


def _category_match_score(raw_name: Any, category: Any) -> float:
    inferred = _infer_product_family(raw_name)
    if not inferred:
        return 0.0
    family = _category_family(category)
    if not family:
        return 0.0
    if inferred == family:
        return 1.0
    if inferred in family or family in inferred:
        return 0.75
    if inferred.endswith("盖") and family.endswith("盖"):
        return 0.55
    if inferred.endswith("杯") and family.endswith("杯"):
        return 0.55
    if inferred in {"吸管", "可降解吸管"} and family in {"吸管", "可降解吸管"}:
        return 0.55
    if inferred == "袋" and family.endswith("袋"):
        return 0.45
    return -0.35


def _product_similarity_score(raw_name: Any, product_name: Any) -> float:
    raw_key = _normalize_product_key(raw_name)
    product_key = _normalize_product_key(product_name)
    if not raw_key or not product_key:
        return 0.0
    if raw_key == product_key:
        return 1.0

    sequence_score = SequenceMatcher(None, raw_key, product_key).ratio()
    raw_chars = set(raw_key)
    product_chars = set(product_key)
    char_overlap = len(raw_chars.intersection(product_chars)) / max(len(raw_chars), 1)
    containment = 1.0 if raw_key in product_key or product_key in raw_key else 0.0
    score = (sequence_score * 0.65) + (char_overlap * 0.25) + (containment * 0.10)

    raw_codes = _product_numeric_codes(raw_name)
    product_codes = _product_numeric_codes(product_name)
    if raw_codes and (
        raw_codes.intersection(product_codes)
        or any(code in product_key for code in raw_codes)
    ):
        score = max(score, 0.92)

    return round(min(score, 1.0), 3)


def _product_match_score(raw_name: Any, product_entry: dict[str, Any]) -> float:
    product_name = product_entry.get("name")
    base_score = _product_similarity_score(raw_name, product_name)
    category_score = _category_match_score(raw_name, product_entry.get("category"))
    if category_score > 0:
        base_score += category_score * 0.08
    elif category_score < 0:
        base_score += category_score * 0.18
    return round(max(0.0, min(base_score, 1.0)), 3)


def _load_product_aliases() -> dict[str, str]:
    aliases: dict[str, str] = {
        _normalize_product_key(key): value
        for key, value in DEFAULT_PRODUCT_ALIASES.items()
        if _normalize_product_key(key) and _normalize_value(value)
    }
    payload = _load_json_file(PRODUCT_ALIAS_FILE)
    if isinstance(payload, dict):
        for key, value in payload.items():
            normalized_key = _normalize_product_key(key)
            normalized_value = _normalize_value(value)
            if normalized_key and normalized_value:
                aliases[normalized_key] = str(normalized_value)
    return aliases


def _load_product_catalog_cache() -> dict[str, Any]:
    payload = _load_json_file(PRODUCT_CACHE_FILE)
    return payload if isinstance(payload, dict) else {}


def _save_product_catalog_cache(payload: dict[str, Any]) -> None:
    os.makedirs(os.path.dirname(PRODUCT_CACHE_FILE) or ".", exist_ok=True)
    _write_json_file(PRODUCT_CACHE_FILE, payload)


def _get_drive_item_metadata(
    file_path: str,
    token: str,
) -> tuple[dict[str, Any] | None, dict[str, Any]]:
    headers = _graph_headers(token)
    proxies = {"http": None, "https": None}
    last_response: requests.Response | None = None
    for url in _build_drive_item_urls(file_path):
        resp = requests.get(
            url,
            headers=headers,
            params={"$select": "id,name,eTag,cTag,lastModifiedDateTime,size,webUrl"},
            proxies=proxies,
            timeout=30,
        )
        last_response = resp
        if resp.status_code == 200:
            return resp.json(), {"success": True, "status_code": resp.status_code}

    status_code = last_response.status_code if last_response is not None else 500
    detail = last_response.text if last_response is not None else "metadata request failed"
    error_type, message = _classify_graph_error(status_code, detail, operation="columns")
    return None, {
        "success": False,
        "error_type": error_type,
        "message": message,
        "status_code": status_code,
        "detail": detail,
    }


def _parse_range_rows(address: str | None) -> tuple[int, int] | None:
    normalized = _normalize_value(address)
    if normalized is None:
        return None
    match = re.search(r"!\$?[A-Z]+\$?(?P<start>\d+)(?::\$?[A-Z]+\$?(?P<end>\d+))?", str(normalized))
    if not match:
        return None
    start_row = int(match.group("start"))
    end_row = int(match.group("end") or start_row)
    return start_row, end_row


def _get_worksheet_used_range(
    workbook_base_url: str,
    sheet_name: str,
    headers: dict[str, str],
    proxies: dict[str, str | None],
) -> tuple[dict[str, Any] | None, dict[str, Any]]:
    worksheet_base = f"{workbook_base_url}/worksheets/{quote(sheet_name, safe='')}"
    last_response: requests.Response | None = None
    for suffix in ("usedRange()", "usedRange"):
        resp = requests.get(f"{worksheet_base}/{suffix}", headers=headers, proxies=proxies, timeout=30)
        last_response = resp
        if resp.status_code == 200:
            return resp.json(), {"success": True, "status_code": resp.status_code}

    status_code = last_response.status_code if last_response is not None else 500
    detail = last_response.text if last_response is not None else "usedRange request failed"
    error_type, message = _classify_graph_error(status_code, detail, operation="rows")
    return None, {
        "success": False,
        "error_type": error_type,
        "message": message,
        "status_code": status_code,
        "detail": detail,
    }


def _load_product_catalog_from_onedrive(
    token: str,
    metadata: dict[str, Any] | None = None,
) -> tuple[dict[str, Any] | None, dict[str, Any]]:
    product_file_path = _normalize_onedrive_path(PRODUCT_FILE_PATH)
    product_sheet_name = _to_string(PRODUCT_SHEET_NAME)
    product_column = _to_string(PRODUCT_NAME_COLUMN).upper()
    category_column = _to_string(PRODUCT_CATEGORY_COLUMN).upper()
    if (
        not product_file_path
        or not product_sheet_name
        or not re.fullmatch(r"[A-Z]+", product_column)
        or not re.fullmatch(r"[A-Z]+", category_column)
    ):
        return None, {
            "success": False,
            "error_type": "invalid_product_catalog_config",
            "message": "产品库配置无效，请检查 OC_OD_PRODUCT_FILE_PATH、OC_OD_PRODUCT_SHEET_NAME、OC_OD_PRODUCT_NAME_COLUMN 和 OC_OD_PRODUCT_CATEGORY_COLUMN。",
        }

    headers = _graph_headers(token)
    proxies = {"http": None, "https": None}
    workbook_base_url = _build_workbook_base_url(product_file_path)
    used_range, used_status = _get_worksheet_used_range(workbook_base_url, product_sheet_name, headers, proxies)
    if not used_range:
        return None, used_status

    parsed_rows = _parse_range_rows(used_range.get("address"))
    if parsed_rows is None:
        return None, {
            "success": False,
            "error_type": "invalid_used_range",
            "message": "无法识别产品明细工作表的 usedRange 地址。",
            "address": used_range.get("address"),
        }

    start_row, end_row = parsed_rows
    if end_row < start_row:
        end_row = start_row
    product_col_index = _column_letters_to_index(product_column)
    category_col_index = _column_letters_to_index(category_column)
    first_col_index = min(product_col_index, category_col_index)
    last_col_index = max(product_col_index, category_col_index)
    first_col = _column_index_to_letter(first_col_index)
    last_col = _column_index_to_letter(last_col_index)
    product_offset = product_col_index - first_col_index
    category_offset = category_col_index - first_col_index
    range_address = f"{first_col}{start_row}:{last_col}{end_row}"
    resp_range = requests.get(
        f"{workbook_base_url}/worksheets/{quote(product_sheet_name, safe='')}/range(address='{range_address}')",
        headers=headers,
        proxies=proxies,
        timeout=30,
    )
    if resp_range.status_code != 200:
        error_type, message = _classify_graph_error(resp_range.status_code, resp_range.text, operation="rows")
        return None, {
            "success": False,
            "error_type": error_type,
            "message": message,
            "status_code": resp_range.status_code,
            "detail": resp_range.text,
            "range_address": range_address,
        }

    products: list[str] = []
    entries: list[dict[str, str | None]] = []
    category_counts: dict[str, int] = {}
    seen: set[str] = set()
    header_names = {"产品名称", "货品名称", "商品名称", "名称", "品名"}
    for row in resp_range.json().get("values", []):
        value = _normalize_value(row[product_offset] if len(row) > product_offset else None)
        if value is None:
            continue
        product_name = str(value)
        if product_name in header_names:
            continue
        category = _normalize_entity_name(str(row[category_offset])) if len(row) > category_offset else None
        product_key = _normalize_product_key(product_name)
        if not product_key or product_key in seen:
            continue
        seen.add(product_key)
        products.append(product_name)
        entries.append({"name": product_name, "category": category})
        category_key = category or "未分类"
        category_counts[category_key] = category_counts.get(category_key, 0) + 1

    post_read_metadata, _ = _get_drive_item_metadata(product_file_path, token)
    source_metadata = post_read_metadata or metadata or {}
    payload = {
        "file_path": product_file_path,
        "sheet_name": product_sheet_name,
        "name_column": product_column,
        "category_column": category_column,
        "source_id": source_metadata.get("id"),
        "source_etag": source_metadata.get("eTag"),
        "source_ctag": source_metadata.get("cTag"),
        "source_last_modified": source_metadata.get("lastModifiedDateTime"),
        "loaded_at": datetime.now().isoformat(timespec="seconds"),
        "product_count": len(products),
        "products": products,
        "entries": entries,
        "category_counts": category_counts,
    }
    _save_product_catalog_cache(payload)
    return payload, {
        "success": True,
        "status": "refreshed",
        "product_count": len(products),
        "range_address": range_address,
    }


def _ensure_product_catalog_fresh(
    token: str | None = None,
    force_refresh: bool = False,
) -> tuple[dict[str, Any], dict[str, Any]]:
    cache = _load_product_catalog_cache()
    if token is None:
        token = get_token_automatically()
    if not token:
        if cache.get("products"):
            return cache, {
                "success": True,
                "status": "stale_cache",
                "message": "无法获取 Microsoft Token，已使用本地产品缓存。",
                "product_count": len(cache.get("products") or []),
            }
        return {}, {
            "success": False,
            "status": "unavailable",
            "error_type": "auth_failed",
            "message": "无法获取 Microsoft Token，且本地产品缓存不存在。",
        }

    metadata, metadata_status = _get_drive_item_metadata(PRODUCT_FILE_PATH, token)
    if metadata is None:
        if cache.get("products"):
            return cache, {
                "success": True,
                "status": "stale_cache",
                "message": "无法读取 OneDrive 产品库元数据，已使用本地产品缓存。",
                "metadata_error": metadata_status,
                "product_count": len(cache.get("products") or []),
            }
        return {}, metadata_status

    same_source = (
        not force_refresh
        and cache.get("products")
        and cache.get("source_etag") == metadata.get("eTag")
        and cache.get("source_last_modified") == metadata.get("lastModifiedDateTime")
    )
    if same_source:
        return cache, {
            "success": True,
            "status": "cache_hit",
            "product_count": len(cache.get("products") or []),
            "source_etag": metadata.get("eTag"),
            "source_last_modified": metadata.get("lastModifiedDateTime"),
        }

    return _load_product_catalog_from_onedrive(token, metadata=metadata)


def _product_catalog_entries(catalog: dict[str, Any] | None) -> list[dict[str, str | None]]:
    if not catalog:
        return []
    entries = catalog.get("entries")
    if isinstance(entries, list):
        normalized_entries: list[dict[str, str | None]] = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            name = _normalize_value(entry.get("name"))
            if name is None:
                continue
            normalized_entries.append(
                {
                    "name": str(name),
                    "category": _normalize_entity_name(entry.get("category")),
                }
            )
        if normalized_entries:
            return normalized_entries

    return [
        {"name": str(product), "category": None}
        for product in (catalog.get("products") or [])
        if _normalize_value(product) is not None
    ]


def _classify_product_name_pattern(product_name: Any, category: Any = None) -> str:
    name = _to_string(product_name)
    family = _category_family(category)
    if re.match(r"^\d{2,3}-[^-]+-[^-]+", name):
        if family.endswith("杯") or "ml" in _normalize_product_key(name) or "oz" in _normalize_product_key(name):
            return "口径-容量/规格-容量/颜色"
        return "数字规格-属性-属性"
    if re.match(r"^\d{2,3}.+盖", name):
        return "口径+盖型"
    if re.match(r"^\d{3}.+管", name):
        return "长度+包装/材质+管型"
    if re.search(r"\d+\*\d+", name):
        return "克重/材质+尺寸"
    if "费" in name or "定金" in name or "优惠" in name:
        return "费用项"
    return "名称描述型"


def _analyze_product_catalog_patterns(catalog: dict[str, Any]) -> list[dict[str, Any]]:
    entries = _product_catalog_entries(catalog)
    grouped: dict[str, list[dict[str, str | None]]] = {}
    for entry in entries:
        category = entry.get("category") or "未分类"
        grouped.setdefault(category, []).append(entry)

    summaries: list[dict[str, Any]] = []
    for category, category_entries in sorted(grouped.items(), key=lambda item: (-len(item[1]), item[0])):
        pattern_counts: dict[str, int] = {}
        samples_by_pattern: dict[str, list[str]] = {}
        for entry in category_entries:
            pattern = _classify_product_name_pattern(entry.get("name"), category)
            pattern_counts[pattern] = pattern_counts.get(pattern, 0) + 1
            samples_by_pattern.setdefault(pattern, [])
            if len(samples_by_pattern[pattern]) < 5:
                samples_by_pattern[pattern].append(_to_string(entry.get("name")))
        primary_pattern = max(pattern_counts.items(), key=lambda item: item[1])[0] if pattern_counts else None
        summaries.append(
            {
                "category": category,
                "family": _category_family(category) or None,
                "count": len(category_entries),
                "primary_pattern": primary_pattern,
                "pattern_counts": pattern_counts,
                "samples": samples_by_pattern.get(primary_pattern or "", []),
            }
        )
    return summaries


def _resolve_product_name_from_catalog(
    raw_name: Any,
    catalog: dict[str, Any] | None = None,
    aliases: dict[str, str] | None = None,
) -> dict[str, Any]:
    original_name = _normalize_value(raw_name)
    if original_name is None:
        return {
            "raw_name": None,
            "resolved_name": None,
            "matched": False,
            "needs_review": False,
            "method": "empty",
            "confidence": 0,
        }

    raw_text = str(original_name)
    raw_key = _normalize_product_key(raw_text)
    entries = _product_catalog_entries(catalog)
    products = [entry["name"] for entry in entries if _normalize_value(entry.get("name")) is not None]
    entry_by_name = {entry["name"]: entry for entry in entries}
    product_keys = {_normalize_product_key(product): product for product in products}
    product_key_set = set(product_keys)
    aliases = aliases or _load_product_aliases()

    def _build_result(
        resolved_name: str | None,
        matched: bool,
        method: str,
        confidence: float,
        candidates: list[str] | None = None,
        needs_review: bool = False,
    ) -> dict[str, Any]:
        return {
            "raw_name": raw_text,
            "resolved_name": resolved_name or raw_text,
            "matched": matched,
            "changed": bool(resolved_name and resolved_name != raw_text),
            "method": method,
            "confidence": confidence,
            "candidates": candidates or [],
            "needs_review": needs_review,
            "inferred_category": _infer_product_family(raw_text) or None,
            "matched_category": entry_by_name.get(resolved_name or "", {}).get("category") if resolved_name else None,
        }

    alias_target = aliases.get(raw_key)
    if alias_target:
        alias_target_key = _normalize_product_key(alias_target)
        if bool(products) and alias_target_key in product_key_set:
            return _build_result(product_keys[alias_target_key], True, "alias", 1.0)
        if not products:
            return _build_result(None, False, "alias_unverified_no_catalog", 0.5, [alias_target], needs_review=True)

    if raw_key in product_keys:
        return _build_result(product_keys[raw_key], True, "exact", 1.0)

    raw_codes = _product_numeric_codes(raw_text)
    if raw_codes and products:
        candidates = sorted(
            (
                (_product_match_score(raw_text, entry), entry["name"])
                for entry in entries
                if raw_codes.intersection(_product_numeric_codes(entry["name"]))
                or any(code in _normalize_product_key(entry["name"]) for code in raw_codes)
            ),
            reverse=True,
        )
        unique_candidates = []
        seen_candidate_keys: set[str] = set()
        for _, product in candidates:
            product_key = _normalize_product_key(product)
            if product_key in seen_candidate_keys:
                continue
            seen_candidate_keys.add(product_key)
            unique_candidates.append(product)
        if len(unique_candidates) == 1:
            return _build_result(unique_candidates[0], True, "numeric_code", 0.92)
        if len(unique_candidates) > 1:
            best_score = candidates[0][0]
            second_score = candidates[1][0] if len(candidates) > 1 else 0
            if best_score >= 0.86 and best_score - second_score >= 0.06:
                return _build_result(candidates[0][1], True, "numeric_code_best", best_score)
            return _build_result(None, False, "ambiguous_numeric_code", best_score, unique_candidates[:5], needs_review=True)

    if products:
        scored = sorted(
            (
                (_product_match_score(raw_text, entry), entry["name"])
                for entry in entries
            ),
            reverse=True,
        )
        if scored:
            best_score, best_product = scored[0]
            second_score = scored[1][0] if len(scored) > 1 else 0
            if best_score >= 0.82 and best_score - second_score >= 0.06:
                return _build_result(best_product, True, "closest_catalog_match", best_score)
            if best_score >= 0.70:
                return _build_result(
                    None,
                    False,
                    "ambiguous_catalog_match",
                    best_score,
                    [product for _, product in scored[:5]],
                    needs_review=True,
                )

    return _build_result(None, False, "not_found", 0, needs_review=bool(products))


def _extract_graph_error_code(detail_text: str) -> str | None:
    try:
        payload = json.loads(detail_text)
    except json.JSONDecodeError:
        return None

    error = payload.get("error")
    if isinstance(error, dict):
        code = error.get("code")
        if isinstance(code, str):
            return code
        inner_error = error.get("innerError")
        if isinstance(inner_error, dict):
            inner_code = inner_error.get("code")
            if isinstance(inner_code, str):
                return inner_code
    return None


def _classify_graph_error(status_code: int, detail_text: str, operation: str) -> tuple[str, str]:
    detail_lower = detail_text.lower()
    error_code = (_extract_graph_error_code(detail_text) or "").lower()

    if status_code in (401,):
        return "auth_failed", "登录已失效或未完成授权，请重新登录 Microsoft 账号。"
    if status_code == 403:
        return "permission_denied", "当前账号没有访问该 Excel 文件或表格的权限。"
    if status_code == 404:
        if "itemnotfound" in detail_lower or error_code == "itemnotfound":
            if operation == "columns":
                return "file_not_found", "找不到目标 Excel 文件，请检查 OC_OD_FILE_PATH 是否正确。"
            if operation == "rows":
                return "table_not_found", "找不到目标 Excel 表格，请检查 OC_OD_TABLE_NAME 是否为表格对象名。"
            return "resource_not_found", "找不到目标资源，请检查文件路径和表格名称。"
        if "resourcenotfound" in detail_lower or error_code == "resourcenotfound":
            return "resource_not_found", "找不到目标资源，请检查文件路径和表格名称。"
    if status_code == 400:
        if "invalidversion" in detail_lower or "invalidrequest" in detail_lower or error_code in {"invalidrequest", "badrequest"}:
            return "table_not_found", "Excel 表格名称无效，请确认 OC_OD_TABLE_NAME 是表格对象名而不是工作表名。"
        if "spo license" in detail_lower:
            return "permission_denied", "当前账号或租户没有可用的 OneDrive/SharePoint 许可。"
    return "graph_request_failed", "访问 Microsoft Graph 失败，请检查登录状态、文件路径、表格名称和权限设置。"


class ExcelOrder(BaseModel):
    备注: str | None = None
    发货厂家: str | None = None
    产品供应商: str | None = None
    日期: str | None = None
    单号: str | None = None
    匹配客户别名: str | None = None
    销售员: str | None = None
    客户: str | None = None
    货品名称: str | None = None
    数量: str | int | float | None = None
    数量单位: str | None = None
    销售单价: str | int | float | None = None
    销售金额: str | int | float | None = None
    成本单价: str | int | float | None = None
    成本金额: str | int | float | None = None
    运费: str | int | float | None = None
    利润: str | int | float | None = None
    总货款: str | int | float | None = None
    已收: str | int | float | None = None
    未收: str | int | float | None = None
    收货联系人: str | None = None
    收货人电话: str | None = None
    收货地址: str | None = None
    items: list[OrderItem] = Field(default_factory=list)
    extra_fields: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="after")
    def validate_match_keys(self) -> "ExcelOrder":
        if not _normalize_value(self.单号) and not _normalize_value(self.客户):
            raise ValueError("至少提供 `单号` 或 `客户`。")
        return self

    def finalize(self) -> "ExcelOrder":
        if self.items:
            aggregated = _aggregate_items_for_excel(self.items)
            if _normalize_value(self.货品名称) is None:
                self.货品名称 = _normalize_value(aggregated.get("货品名称"))
            if _normalize_value(self.数量) is None:
                self.数量 = _normalize_value(aggregated.get("数量"))
            if _normalize_value(self.数量单位) is None:
                self.数量单位 = _normalize_value(aggregated.get("数量单位"))
            if _normalize_value(self.销售单价) is None:
                self.销售单价 = _normalize_value(aggregated.get("销售单价"))
            if _normalize_value(self.销售金额) is None:
                self.销售金额 = _normalize_value(aggregated.get("销售金额"))
            self.extra_fields["商品明细"] = json.dumps(
                aggregated.get("items", []),
                ensure_ascii=False,
            )

        total_payment = _to_float(self.总货款)
        received_payment = _to_float(self.已收)
        sales_amount = _to_float(self.销售金额)
        shipping_fee = _to_float(self.运费)
        cost_amount = _to_float(self.成本金额)

        if total_payment is None and sales_amount is not None:
            total_payment = sales_amount
        unpaid = None
        if total_payment is not None and received_payment is not None:
            unpaid = total_payment - received_payment

        if sales_amount is None and total_payment is not None:
            sales_amount = total_payment

        profit = _to_float(self.利润)

        self.日期 = _normalize_date(self.日期, None) or datetime.now().strftime("%Y-%m-%d")
        self.销售金额 = _format_money(sales_amount)
        self.总货款 = _format_money(total_payment)
        self.已收 = _format_money(received_payment)
        self.未收 = _format_money(unpaid)
        self.利润 = _format_money(profit)
        self.销售单价 = _format_unit_price(self.销售单价)
        self.成本单价 = _format_money(_to_float(self.成本单价))
        self.成本金额 = _format_money(cost_amount)
        self.运费 = _format_money(shipping_fee)
        self.数量 = _normalize_value(self.数量)
        return self

    def to_excel_dict(self) -> dict[str, Any]:
        self.finalize()
        base_fields = {header: _normalize_value(getattr(self, header)) for header in EXCEL_HEADERS}
        base_fields.update(
            {
                str(key).strip(): _normalize_value(value)
                for key, value in self.extra_fields.items()
                if str(key).strip()
            }
        )
        return {key: value for key, value in base_fields.items() if value is not None}


class ParsedWechatOrder(BaseModel):
    raw_message: str
    sender_name: str | None = None
    group_name: str | None = None
    message_time: str | None = None
    order: ExcelOrder
    missing_fields: list[str] = Field(default_factory=list)
    needs_review: bool = False


class OrderIngestRequest(BaseModel):
    raw_message: str = Field(description="当前收到的订单相关文本。可以是完整订单，也可以是补充信息。")
    sender_name: str | None = Field(default=None, description="微信群实际发送人。优先用于销售员和匹配约束。")
    group_name: str | None = Field(default=None, description="微信群名称，可选。")
    message_time: str | None = Field(default=None, description="消息时间，可选。")
    existing_order: ExcelOrder | None = Field(
        default=None,
        description="如果当前消息是对已有草稿的补充，这里传入已有草稿订单；否则留空。",
    )
    auto_add_new_columns: bool = Field(
        default=False,
        description="写入 Excel 时是否自动新增新列。默认关闭，避免误加表头。",
    )
    dry_run: bool = Field(
        default=False,
        description="仅解析和匹配，不真正写入 Excel。",
    )


def _standardize_order_products(order: ExcelOrder, token: str | None = None) -> dict[str, Any]:
    if not _has_item_details(order):
        return {
            "enabled": True,
            "attempted": False,
            "needs_review": False,
            "items": [],
        }

    catalog, catalog_status = _ensure_product_catalog_fresh(token=token)
    aliases = _load_product_aliases()
    resolution_items: list[dict[str, Any]] = []

    if order.items:
        standardized_items: list[OrderItem] = []
        for item in order.items:
            if _is_plate_fee_name(item.货品名称):
                resolution_items.append(
                    {
                        "raw_name": item.货品名称,
                        "resolved_name": "版费",
                        "matched": True,
                        "changed": item.货品名称 != "版费",
                        "method": "fee_item",
                        "confidence": 1.0,
                        "candidates": [],
                        "needs_review": False,
                        "inferred_category": "费用",
                        "matched_category": "费用",
                    }
                )
                standardized_items.append(
                    OrderItem(
                        货品名称="版费",
                        数量=item.数量 or "1",
                        数量单位=item.数量单位 or "项",
                        销售单价=item.销售单价,
                        销售金额=item.销售金额,
                    ).normalized()
                )
                continue

            resolved = _resolve_product_name_from_catalog(item.货品名称, catalog=catalog, aliases=aliases)
            resolution_items.append(resolved)
            standardized_items.append(
                OrderItem(
                    货品名称=_to_string(resolved.get("resolved_name")) or item.货品名称,
                    数量=item.数量,
                    数量单位=item.数量单位,
                    销售单价=item.销售单价,
                    销售金额=item.销售金额,
                ).normalized()
            )
        order.items = standardized_items
        order.货品名称 = None
        order.数量 = None
        order.数量单位 = None
        order.销售单价 = None
    elif _normalize_value(order.货品名称) is not None:
        if _is_plate_fee_name(order.货品名称):
            resolved = {
                "raw_name": order.货品名称,
                "resolved_name": "版费",
                "matched": True,
                "changed": order.货品名称 != "版费",
                "method": "fee_item",
                "confidence": 1.0,
                "candidates": [],
                "needs_review": False,
                "inferred_category": "费用",
                "matched_category": "费用",
            }
        else:
            resolved = _resolve_product_name_from_catalog(order.货品名称, catalog=catalog, aliases=aliases)
        resolution_items.append(resolved)
        order.货品名称 = _to_string(resolved.get("resolved_name")) or order.货品名称

    needs_review = any(bool(item.get("needs_review")) for item in resolution_items)
    if resolution_items and not catalog.get("products"):
        needs_review = True
    changed_items = [item for item in resolution_items if item.get("changed") or item.get("needs_review")]
    if changed_items:
        order.extra_fields["商品名称标准化"] = json.dumps(changed_items, ensure_ascii=False)
    if needs_review:
        order.extra_fields["商品名称需复核"] = "是"
    else:
        order.extra_fields.pop("商品名称需复核", None)

    return {
        "enabled": True,
        "attempted": True,
        "needs_review": needs_review,
        "catalog_status": catalog_status,
        "product_count": len(catalog.get("products") or []),
        "items": resolution_items,
    }


def _parse_wechat_order_message_model(
    raw_message: str,
    sender_name: str | None = None,
    group_name: str | None = None,
    message_time: str | None = None,
) -> ParsedWechatOrder:
    lines = [line.strip() for line in raw_message.splitlines() if line.strip()]

    order_number, customer_alias = _extract_order_number_and_customer_alias(raw_message)
    contact_name = _extract_contact_name(raw_message)
    phone = _extract_phone(raw_message)
    address = _extract_address(raw_message)
    explicit_salesperson = _extract_salesperson_name(raw_message)
    order_date = _normalize_date(order_number, message_time)
    total_amount = _extract_first(r"(?:全款|总货款|合计|共计|共)[:：]?[ \t]*([0-9]+(?:\.[0-9]+)?)\s*元?", raw_message)
    received_amount = _extract_first(
        r"(?:已收定金|已付定金|收定金|定金|已收|实收|收到)[:：]?[ \t]*([0-9]+(?:\.[0-9]+)?)\s*元?",
        raw_message,
    )
    payment_note = _extract_first(r"(.+收款)", raw_message)
    item_count = _extract_first(r"(?:商品|货品|产品)总数[:：]\s*(\d+)", raw_message)

    items = _extract_order_items(raw_message)
    aggregated_items = _aggregate_items_for_excel(items)
    message_intent = _detect_message_intent(raw_message, bool(items))
    if total_amount is None:
        total_amount = _normalize_value(aggregated_items.get("销售金额"))
    if received_amount is None and re.search(r"全款\s*\d+(?:\.\d+)?\s*元?", raw_message):
        received_amount = total_amount

    explicit_customer = _extract_first(r"客户[:：]\s*([^\r\n]+)", raw_message)
    customer = _normalize_replace_target(explicit_customer) if explicit_customer else None
    if customer is None:
        customer = customer_alias
    if customer is None:
        customer = _extract_customer_name(raw_message)
    if customer is None:
        customer = contact_name

    order = ExcelOrder(
        日期=order_date,
        单号=order_number,
        匹配客户别名=_normalize_value(customer_alias),
        销售员=_normalize_value(explicit_salesperson or sender_name),
        客户=_normalize_value(customer),
        货品名称=_normalize_value(aggregated_items.get("货品名称")),
        数量=_normalize_value(aggregated_items.get("数量")),
        数量单位=_normalize_value(aggregated_items.get("数量单位")),
        销售单价=_normalize_value(aggregated_items.get("销售单价")),
        销售金额=_normalize_value(total_amount),
        总货款=_normalize_value(total_amount),
        已收=_normalize_value(received_amount),
        收货联系人=_normalize_value(contact_name or customer),
        收货人电话=_normalize_value(phone),
        收货地址=_normalize_value(address),
        备注=_normalize_value(payment_note),
        items=items,
        extra_fields=(
            {
                "商品总数": item_count,
                "商品明细": json.dumps(aggregated_items.get("items", []), ensure_ascii=False)
                if aggregated_items.get("items")
                else None,
                "消息意图": message_intent,
            }
        ),
    )
    order.extra_fields = {
        key: value for key, value in order.extra_fields.items() if _normalize_value(value) is not None
    }

    missing_fields = [
        field_name
        for field_name in ("单号", "销售员", "客户", "货品名称", "数量", "销售金额", "收货人电话", "收货地址")
        if _normalize_value(getattr(order, field_name)) is None
    ]
    return ParsedWechatOrder(
        raw_message=raw_message,
        sender_name=sender_name,
        group_name=group_name,
        message_time=message_time,
        order=order,
        missing_fields=missing_fields,
        needs_review=bool(missing_fields),
    )


def _merge_orders(existing_order: ExcelOrder, new_order: ExcelOrder, sender_name: str | None = None) -> ExcelOrder:
    merged = existing_order.model_copy(deep=True)

    merged.备注 = _merge_prefer_new(existing_order.备注, new_order.备注)
    merged.发货厂家 = _merge_prefer_new(existing_order.发货厂家, new_order.发货厂家)
    merged.产品供应商 = _merge_prefer_new(existing_order.产品供应商, new_order.产品供应商)
    merged.日期 = _merge_prefer_new(existing_order.日期, new_order.日期)
    merged.单号 = _merge_prefer_new(existing_order.单号, new_order.单号)
    merged.匹配客户别名 = _merge_prefer_new(existing_order.匹配客户别名, new_order.匹配客户别名)
    merged.销售员 = _merge_prefer_new(existing_order.销售员, sender_name or new_order.销售员)
    merged.客户 = _merge_prefer_new(existing_order.客户, new_order.客户)
    merged.货品名称 = _merge_prefer_new(existing_order.货品名称, new_order.货品名称)
    merged.数量 = _merge_prefer_new(existing_order.数量, new_order.数量)
    merged.数量单位 = _merge_prefer_new(existing_order.数量单位, new_order.数量单位)
    merged.销售单价 = _merge_prefer_new(existing_order.销售单价, new_order.销售单价)
    merged.销售金额 = _merge_prefer_new(existing_order.销售金额, new_order.销售金额)
    merged.成本单价 = _merge_prefer_new(existing_order.成本单价, new_order.成本单价)
    merged.成本金额 = _merge_prefer_new(existing_order.成本金额, new_order.成本金额)
    merged.运费 = _merge_prefer_new(existing_order.运费, new_order.运费)
    merged.利润 = _merge_prefer_new(existing_order.利润, new_order.利润)
    merged.总货款 = _merge_prefer_new(existing_order.总货款, new_order.总货款)
    merged.已收 = _merge_prefer_new(existing_order.已收, new_order.已收)
    merged.未收 = _merge_prefer_new(existing_order.未收, new_order.未收)
    merged.收货联系人 = _merge_prefer_new(existing_order.收货联系人, new_order.收货联系人)
    merged.收货人电话 = _merge_prefer_new(existing_order.收货人电话, new_order.收货人电话)
    merged.收货地址 = _merge_prefer_new(existing_order.收货地址, new_order.收货地址)
    merged.items = new_order.items or existing_order.items

    merged_extra = dict(existing_order.extra_fields)
    for key, value in new_order.extra_fields.items():
        normalized_key = str(key).strip()
        if normalized_key:
            merged_extra[normalized_key] = _merge_prefer_new(merged_extra.get(normalized_key), value)
    merged.extra_fields = merged_extra
    return merged


def _build_order_from_matched_rows(
    existing_columns: list[str],
    rows_data: list[dict[str, Any]],
    matched_row_indexes: list[int],
) -> ExcelOrder:
    if not matched_row_indexes:
        raise ValueError("matched_row_indexes 不能为空")

    base_fields: dict[str, Any] = {}
    items: list[OrderItem] = []

    for row_index in matched_row_indexes:
        row_values = rows_data[row_index].get("values", [[]])[0]
        row_map = {
            column_name: (row_values[idx] if idx < len(row_values) else "")
            for idx, column_name in enumerate(existing_columns)
        }

        for field_name in EXCEL_HEADERS:
            if field_name in {"货品名称", "数量", "数量单位", "销售单价", "销售金额", "总货款", "已收", "未收"}:
                continue
            if _normalize_value(base_fields.get(field_name)) is None:
                base_fields[field_name] = _normalize_value(row_map.get(field_name))

        item_name = _normalize_value(row_map.get("货品名称"))
        item_qty = _normalize_value(row_map.get("数量"))
        item_unit = _normalize_value(row_map.get("数量单位"))
        item_unit_price = _normalize_value(row_map.get("销售单价"))
        item_amount = _normalize_value(row_map.get("销售金额"))
        if any(value is not None for value in (item_name, item_qty, item_unit, item_unit_price, item_amount)):
            items.append(
                OrderItem(
                    货品名称=_to_string(item_name) or None,
                    数量=_to_string(item_qty) or None,
                    数量单位=_to_string(item_unit) or None,
                    销售单价=_to_string(item_unit_price) or None,
                    销售金额=_to_string(item_amount) or None,
                )
            )

    first_row_values = rows_data[matched_row_indexes[0]].get("values", [[]])[0]
    first_row_map = {
        column_name: (first_row_values[idx] if idx < len(first_row_values) else "")
        for idx, column_name in enumerate(existing_columns)
    }
    base_fields["已收"] = _normalize_value(first_row_map.get("已收"))
    base_fields["未收"] = _normalize_value(first_row_map.get("未收"))

    numeric_compatible_fields = {
        "数量",
        "销售单价",
        "销售金额",
        "成本单价",
        "成本金额",
        "运费",
        "利润",
        "总货款",
        "已收",
        "未收",
    }
    order_kwargs: dict[str, Any] = {}
    for field_name in EXCEL_HEADERS:
        value = _normalize_value(base_fields.get(field_name))
        if value is None:
            order_kwargs[field_name] = None
        elif field_name in numeric_compatible_fields:
            order_kwargs[field_name] = value
        else:
            order_kwargs[field_name] = _to_string(value) or None

    return ExcelOrder(**order_kwargs, items=items)


mcp = FastMCP(
    name="Chengyi_Order_Manager",
    instructions=(
        "Order entry MCP for parsing WeChat order messages, merging follow-up updates, "
        "and writing the final structured order into OneDrive Excel."
    ),
    host=DEFAULT_HOST,
    port=DEFAULT_PORT,
    stateless_http=True,
    json_response=True,
)


@mcp.tool()
def health() -> str:
    """
    MCP 服务快速自检。

    不访问 Microsoft Graph，不触发登录，只用于 OpenClaw 启动后确认服务已加载、
    核心配置是否存在，以及当前监听参数是否生效。
    """
    client_id = os.getenv("OC_OD_CLIENT_ID")
    os.makedirs(LOGS_DIR, exist_ok=True)
    cache_exists = os.path.exists(CACHE_FILE)
    result = _json_result(
        success=True,
        action="health",
        message="MCP 服务可用。",
        service_name="Chengyi_Order_Manager",
        transport_default=os.getenv("CY_EXCEL_MCP_TRANSPORT", "streamable-http"),
        host=DEFAULT_HOST,
        port=DEFAULT_PORT,
        authority=AUTHORITY,
        tenant_id=TENANT_ID,
        file_path=FILE_PATH,
        table_name=TABLE_NAME,
        product_file_path=_normalize_onedrive_path(PRODUCT_FILE_PATH),
        product_sheet_name=PRODUCT_SHEET_NAME,
        product_name_column=PRODUCT_NAME_COLUMN,
        product_category_column=PRODUCT_CATEGORY_COLUMN,
        product_cache_file=PRODUCT_CACHE_FILE,
        product_cache_exists=os.path.exists(PRODUCT_CACHE_FILE),
        product_alias_file=PRODUCT_ALIAS_FILE,
        product_alias_file_exists=os.path.exists(PRODUCT_ALIAS_FILE),
        client_id_configured=bool(client_id),
        cache_file=CACHE_FILE,
        cache_exists=cache_exists,
        logs_dir=LOGS_DIR,
        logs_dir_exists=os.path.isdir(LOGS_DIR),
    )
    _append_tool_log("health", result_type="health")
    return result


@mcp.tool()
def check_login_status() -> str:
    """
    检查当前 Microsoft 登录状态。

    只做只读检查，不会主动触发 device code 登录。
    用于判断当前是否已命中 token cache、是否已有缓存账号、以及是否可以静默获取 token。
    """
    client_id = os.getenv("OC_OD_CLIENT_ID")
    cache_exists = os.path.exists(CACHE_FILE)
    base_info = {
        "authority": AUTHORITY,
        "cache_file": CACHE_FILE,
        "cache_exists": cache_exists,
        "tenant_id": TENANT_ID,
        "file_path": FILE_PATH,
        "table_name": TABLE_NAME,
        "client_id_configured": bool(client_id),
    }
    if not client_id:
        result = _json_result(
            success=False,
            action="login_status",
            message="未配置 OC_OD_CLIENT_ID。",
            **base_info,
        )
        _append_tool_log("check_login_status", result_type="login_status_missing_client_id")
        return result

    cache = _build_token_cache()
    accounts: list[dict[str, Any]] = []
    try:
        app = _build_public_client_application(token_cache=cache)
        if app is None:
            result = _json_result(
                success=False,
                action="login_status",
                message="无法初始化 Microsoft 登录客户端。",
                **base_info,
            )
            _append_tool_log("check_login_status", result_type="login_status_init_failed")
            return result

        accounts = app.get_accounts()
        silent_token_available = False
        silent_check = "not_attempted"
        detail = None
        if accounts:
            try:
                result = app.acquire_token_silent(SCOPES, account=accounts[0])
                silent_token_available = bool(result and "access_token" in result)
                silent_check = "ok"
            except Exception as exc:
                detail = str(exc)
                silent_check = "network_error" if _is_network_error(exc) else "failed"

        result = _json_result(
            success=True,
            action="login_status",
            message="已完成登录状态检查。",
            account_count=len(accounts),
            has_cached_account=bool(accounts),
            has_silent_token=silent_token_available,
            cached_username=_normalize_value(accounts[0].get("username")) if accounts else None,
            silent_check=silent_check,
            needs_login=not bool(accounts),
            needs_network_retry=bool(accounts) and not silent_token_available and silent_check == "network_error",
            detail=detail,
            **base_info,
        )
        _append_tool_log(
            "check_login_status",
            order_number=None,
            result_type="login_status",
            has_cached_account=bool(accounts),
            has_silent_token=silent_token_available,
            cached_username=_normalize_value(accounts[0].get("username")) if accounts else None,
        )
        return result
    except Exception as exc:
        cache_state = {
            "account_count": len(accounts),
            "has_cached_account": bool(accounts),
            "has_silent_token": False,
            "cached_username": _normalize_value(accounts[0].get("username")) if accounts else None,
        }
        if _is_network_error(exc):
            result = _json_result(
                success=True,
                action="login_status",
                message="已读取本地登录缓存，但当前网络不可用，无法完成在线校验。",
                silent_check="network_error",
                needs_login=not bool(accounts),
                needs_network_retry=True,
                detail=str(exc),
                **cache_state,
                **base_info,
            )
            _append_tool_log("check_login_status", result_type="login_status_network_error")
            return result
        result = _json_result(
            success=False,
            action="login_status",
            message="登录状态检查失败。",
            detail=str(exc),
            **cache_state,
            **base_info,
        )
        _append_tool_log("check_login_status", result_type="login_status_failed")
        return result


@mcp.tool()
def list_excel_tables(file_path: str | None = None) -> str:
    """
    列出目标 Excel workbook 中的所有表格对象名称。

    用于辅助确认 `OC_OD_TABLE_NAME` 是否填写正确。
    - `file_path` 留空时使用环境变量 `OC_OD_FILE_PATH`
    - 会访问 Microsoft Graph，必要时可能触发 device flow 登录
    """
    target_file_path = _normalize_value(file_path) or FILE_PATH
    try:
        token = get_token_automatically()
    except Exception as exc:
        result = _json_result(
            success=False,
            action="list_excel_tables",
            error_type="auth_failed",
            message="无法初始化 Microsoft 登录状态，请检查网络或稍后重试。",
            file_path=target_file_path,
            detail=str(exc),
        )
        _append_tool_log("list_excel_tables", result_type="auth_failed", file_path=target_file_path)
        return result
    if not token:
        result = _json_result(
            success=False,
            action="list_excel_tables",
            error_type="auth_failed",
            message="无法获取微软授权 Token，请确认已设置环境变量 OC_OD_CLIENT_ID 并完成设备登录。",
            file_path=target_file_path,
        )
        _append_tool_log("list_excel_tables", result_type="auth_failed")
        return result

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    proxies = {"http": None, "https": None}
    workbook_base_url = _build_workbook_base_url(target_file_path)

    try:
        resp_tables = requests.get(f"{workbook_base_url}/tables", headers=headers, proxies=proxies, timeout=30)
        if resp_tables.status_code != 200:
            error_type, error_message = _classify_graph_error(
                resp_tables.status_code,
                resp_tables.text,
                operation="columns",
            )
            result = _json_result(
                success=False,
                action="list_excel_tables",
                error_type=error_type,
                message=error_message,
                file_path=target_file_path,
                status_code=resp_tables.status_code,
                detail=resp_tables.text,
            )
            _append_tool_log("list_excel_tables", result_type=error_type, file_path=target_file_path)
            return result

        tables = resp_tables.json().get("value", [])
        table_names = [table.get("name") for table in tables if table.get("name")]
        result = _json_result(
            success=True,
            action="list_excel_tables",
            message="已读取 workbook 中的表格对象列表。",
            file_path=target_file_path,
            configured_table_name=TABLE_NAME,
            table_count=len(table_names),
            table_names=table_names,
            table_name_exists=TABLE_NAME in table_names,
        )
        _append_tool_log("list_excel_tables", result_type="success", file_path=target_file_path)
        return result
    except Exception as exc:
        result = _json_result(
            success=False,
            action="list_excel_tables",
            error_type="exception",
            message="读取 workbook 表格对象列表失败。",
            file_path=target_file_path,
            detail=str(exc),
        )
        _append_tool_log("list_excel_tables", result_type="exception", file_path=target_file_path)
        return result


@mcp.tool()
def check_product_catalog_status(check_remote: bool = False) -> str:
    """
    检查本地产品库缓存状态。

    默认只读本地缓存，不访问 Microsoft Graph。
    `check_remote=true` 时会读取 OneDrive 产品库文件元数据，用于判断缓存是否已经过期。
    """
    cache = _load_product_catalog_cache()
    aliases = _load_product_aliases()
    result_payload: dict[str, Any] = {
        "success": True,
        "action": "check_product_catalog_status",
        "file_path": _normalize_onedrive_path(PRODUCT_FILE_PATH),
        "sheet_name": PRODUCT_SHEET_NAME,
        "name_column": PRODUCT_NAME_COLUMN,
        "category_column": PRODUCT_CATEGORY_COLUMN,
        "cache_file": PRODUCT_CACHE_FILE,
        "cache_exists": bool(cache),
        "product_count": len(cache.get("products") or []),
        "category_counts": cache.get("category_counts") or {},
        "alias_file": PRODUCT_ALIAS_FILE,
        "alias_count": len(aliases),
        "source_etag": cache.get("source_etag"),
        "source_last_modified": cache.get("source_last_modified"),
        "loaded_at": cache.get("loaded_at"),
    }
    if check_remote:
        token = get_token_automatically()
        if not token:
            result_payload.update(
                {
                    "success": False,
                    "error_type": "auth_failed",
                    "message": "无法获取 Microsoft Token，不能检查远端产品库元数据。",
                }
            )
        else:
            metadata, metadata_status = _get_drive_item_metadata(PRODUCT_FILE_PATH, token)
            result_payload["remote_check"] = metadata_status
            if metadata:
                result_payload.update(
                    {
                        "remote_etag": metadata.get("eTag"),
                        "remote_last_modified": metadata.get("lastModifiedDateTime"),
                        "remote_changed": (
                            cache.get("source_etag") != metadata.get("eTag")
                            or cache.get("source_last_modified") != metadata.get("lastModifiedDateTime")
                        ),
                    }
                )

    _append_tool_log("check_product_catalog_status", result_type="success")
    return _json_result(**result_payload)


@mcp.tool()
def refresh_product_catalog(include_products: bool = False) -> str:
    """
    强制从 OneDrive 产品库刷新本地产品缓存。

    默认只返回产品数量和缓存元数据，不返回产品列表，避免聊天窗口刷屏。
    """
    try:
        token = get_token_automatically()
    except Exception as exc:
        result = _json_result(
            success=False,
            action="refresh_product_catalog",
            error_type="auth_failed",
            message="无法初始化 Microsoft 登录状态，请检查网络或稍后重试。",
            detail=str(exc),
        )
        _append_tool_log("refresh_product_catalog", result_type="auth_failed")
        return result
    catalog, status = _ensure_product_catalog_fresh(token=token, force_refresh=True)
    products = catalog.get("products") or []
    result_payload = {
        "success": bool(status.get("success")),
        "action": "refresh_product_catalog",
        "message": "产品库缓存已刷新。" if status.get("success") else "产品库缓存刷新失败。",
        "file_path": _normalize_onedrive_path(PRODUCT_FILE_PATH),
        "sheet_name": PRODUCT_SHEET_NAME,
        "name_column": PRODUCT_NAME_COLUMN,
        "category_column": PRODUCT_CATEGORY_COLUMN,
        "cache_file": PRODUCT_CACHE_FILE,
        "product_count": len(products),
        "category_counts": catalog.get("category_counts") or {},
        "catalog_status": status,
        "source_etag": catalog.get("source_etag"),
        "source_last_modified": catalog.get("source_last_modified"),
        "loaded_at": catalog.get("loaded_at"),
    }
    if include_products:
        result_payload["products"] = products
    _append_tool_log("refresh_product_catalog", result_type=status.get("status") or status.get("error_type"))
    return _json_result(**result_payload)


@mcp.tool()
def analyze_product_catalog_patterns(force_refresh: bool = False, include_samples: bool = True) -> str:
    """
    按 C 列分类分析产品明细 B 列命名规律。

    用于排查商品语义匹配规则，不写入订单 Excel。
    """
    try:
        token = get_token_automatically()
    except Exception as exc:
        result = _json_result(
            success=False,
            action="analyze_product_catalog_patterns",
            error_type="auth_failed",
            message="无法初始化 Microsoft 登录状态，请检查网络或稍后重试。",
            detail=str(exc),
        )
        _append_tool_log("analyze_product_catalog_patterns", result_type="auth_failed")
        return result

    catalog, status = _ensure_product_catalog_fresh(token=token, force_refresh=force_refresh)
    summaries = _analyze_product_catalog_patterns(catalog)
    if not include_samples:
        summaries = [
            {key: value for key, value in item.items() if key != "samples"}
            for item in summaries
        ]
    result = _json_result(
        success=bool(status.get("success")),
        action="analyze_product_catalog_patterns",
        message="已按产品分类分析 B 列命名规律。" if status.get("success") else "产品库不可用，无法分析命名规律。",
        file_path=_normalize_onedrive_path(PRODUCT_FILE_PATH),
        sheet_name=PRODUCT_SHEET_NAME,
        name_column=PRODUCT_NAME_COLUMN,
        category_column=PRODUCT_CATEGORY_COLUMN,
        product_count=len(catalog.get("products") or []),
        category_count=len(summaries),
        catalog_status=status,
        patterns=summaries,
    )
    _append_tool_log("analyze_product_catalog_patterns", result_type=status.get("status") or status.get("error_type"))
    return result


@mcp.tool()
def resolve_product_name(raw_name: str, force_refresh: bool = False) -> str:
    """
    将业务员发送的原始商品名解析为产品明细表中的标准产品名称。

    用于单独测试商品名标准化，不写入订单 Excel。
    """
    try:
        token = get_token_automatically()
    except Exception as exc:
        result = _json_result(
            success=False,
            action="resolve_product_name",
            error_type="auth_failed",
            message="无法初始化 Microsoft 登录状态，请检查网络或稍后重试。",
            detail=str(exc),
        )
        _append_tool_log("resolve_product_name", result_type="auth_failed")
        return result

    catalog, status = _ensure_product_catalog_fresh(token=token, force_refresh=force_refresh)
    resolved = _resolve_product_name_from_catalog(raw_name, catalog=catalog, aliases=_load_product_aliases())
    result = _json_result(
        success=True,
        action="resolve_product_name",
        raw_name=raw_name,
        resolved_name=resolved.get("resolved_name"),
        matched=resolved.get("matched"),
        changed=resolved.get("changed"),
        method=resolved.get("method"),
        confidence=resolved.get("confidence"),
        inferred_category=resolved.get("inferred_category"),
        matched_category=resolved.get("matched_category"),
        candidates=resolved.get("candidates"),
        needs_review=resolved.get("needs_review"),
        catalog_status=status,
        product_count=len(catalog.get("products") or []),
    )
    _append_tool_log("resolve_product_name", result_type=resolved.get("method"))
    return result


@mcp.tool()
def parse_wechat_order_message(
    raw_message: str,
    sender_name: str | None = None,
    group_name: str | None = None,
    message_time: str | None = None,
) -> str:
    """
    解析微信群订单消息，输出与你 Excel 表头一致的标准订单对象。

    建议：
    - `sender_name` 直接传微信群实际发送人，优先用它作为 `销售员`
    - `message_time` 传原始消息时间，解析不到 `日期` 时可用它兜底
    - 若字段缺失，会在返回结果里标记 `missing_fields` 和 `needs_review`
    """
    parsed = _parse_wechat_order_message_model(
        raw_message=raw_message,
        sender_name=sender_name,
        group_name=group_name,
        message_time=message_time,
    )
    result = parsed.model_dump_json(indent=2, exclude_none=True, ensure_ascii=False)
    _append_tool_log("parse_wechat_order_message", order_number=parsed.order.单号, result_type="parsed")
    return result


@mcp.tool()
def merge_order_update(
    existing_order: ExcelOrder,
    raw_message: str,
    sender_name: str | None = None,
    group_name: str | None = None,
    message_time: str | None = None,
) -> str:
    """
    合并“图片草稿订单 + 后续补充文字”。

    合并规则：
    - 新消息中的非空字段覆盖草稿中的空字段
    - 若新消息带有重复订单号，则视为同一订单的更新
    - `sender_name` 会优先写入 `销售员`，并作为后续 Excel 倒序匹配时的重要辅助条件
    """
    parsed = _parse_wechat_order_message_model(
        raw_message=raw_message,
        sender_name=sender_name or existing_order.销售员,
        group_name=group_name,
        message_time=message_time,
    )
    merged_order = _merge_orders(
        existing_order=existing_order,
        new_order=parsed.order,
        sender_name=sender_name,
    )

    duplicate_order_number = bool(
        _normalize_value(existing_order.单号)
        and _normalize_value(parsed.order.单号)
        and _same_text(existing_order.单号, parsed.order.单号)
    )

    missing_fields = [
        field_name
        for field_name in ("单号", "销售员", "客户", "货品名称", "数量", "销售金额", "收货人电话", "收货地址")
        if _normalize_value(getattr(merged_order, field_name)) is None
    ]
    result = _json_result(
        success=True,
        action="merged",
        duplicate_order_number=duplicate_order_number,
        message="已将补充文字合并到现有订单草稿。",
        match_hints={
            "单号": _normalize_value(merged_order.单号),
            "匹配客户别名": _normalize_value(merged_order.匹配客户别名),
            "客户": _normalize_value(merged_order.客户),
            "销售员": _normalize_value(merged_order.销售员),
        },
        missing_fields=missing_fields,
        needs_review=bool(missing_fields),
        order=merged_order.to_excel_dict(),
    )
    _append_tool_log("merge_order_update", order_number=merged_order.单号, result_type="merged")
    return result


@mcp.tool()
def ingest_order_message(request: OrderIngestRequest) -> str:
    """
    统一入口：接收一条订单消息，自动完成解析、可选合并，以及写入 Excel。

    调用规则：
    - 纯文本完整订单：只传 `raw_message`
    - 同一订单的补充消息：同时传 `raw_message` 和 `existing_order`
    - `sender_name` 建议始终传，便于提升匹配成功率
    """
    parsed = _parse_wechat_order_message_model(
        raw_message=request.raw_message,
        sender_name=request.sender_name,
        group_name=request.group_name,
        message_time=request.message_time,
    )

    existing_order = request.existing_order
    draft_row_indexes: list[int] = []
    historical_row_indexes: list[int] = []
    if existing_order is None:
        prefer_pending_replace = (
            _to_string(parsed.order.extra_fields.get("消息意图")) == "replace_order"
            and not _has_item_details(parsed.order)
        )
        recent_draft = _find_recent_draft(
            parsed.order,
            request.sender_name,
            prefer_pending_replace=prefer_pending_replace,
        )
        if recent_draft:
            try:
                existing_order = ExcelOrder.model_validate(recent_draft.get("order") or {})
                draft_row_indexes = _normalize_row_indexes(recent_draft.get("row_indexes"))
                historical_row_indexes = _normalize_row_indexes(recent_draft.get("historical_row_indexes"))
            except Exception:
                existing_order = None
                draft_row_indexes = []
                historical_row_indexes = []

    if existing_order is not None:
        final_order = _merge_orders(
            existing_order=existing_order,
            new_order=parsed.order,
            sender_name=request.sender_name,
        )
        if draft_row_indexes:
            final_order.extra_fields["最近草稿行索引"] = json.dumps(draft_row_indexes, ensure_ascii=False)
        if historical_row_indexes:
            final_order.extra_fields["历史订单行索引"] = json.dumps(historical_row_indexes, ensure_ascii=False)
        pipeline_action = "merged_then_processed"
        duplicate_order_number = bool(
            _normalize_value(existing_order.单号)
            and _normalize_value(parsed.order.单号)
            and _same_text(existing_order.单号, parsed.order.单号)
        )
    else:
        final_order = parsed.order
        pipeline_action = "parsed_then_processed"
        duplicate_order_number = False

    process_result = process_excel_order(
        order=final_order,
        auto_add_new_columns=request.auto_add_new_columns,
        dry_run=request.dry_run,
    )
    process_payload = json.loads(process_result)
    should_cache_draft = process_payload.get("success")
    if (
        not should_cache_draft
        and process_payload.get("action") == "needs_review"
        and process_payload.get("error_type") == "row_count_mismatch"
        and _has_item_details(final_order)
    ):
        should_cache_draft = True

    if should_cache_draft and not request.dry_run:
        cached_row_indexes = _normalize_row_indexes(process_payload.get("row_indexes"))
        if not cached_row_indexes:
            cached_row_indexes = _normalize_row_indexes(process_payload.get("replaced_row_indexes"))
        historical_indexes = _normalize_row_indexes(process_payload.get("historical_row_indexes"))
        if not historical_indexes:
            historical_indexes = _normalize_row_indexes(process_payload.get("row_indexes"))
        pending_replace = (
            process_payload.get("action") == "needs_review"
            and process_payload.get("error_type") == "row_count_mismatch"
            and _has_item_details(final_order)
        )
        _store_recent_draft(
            final_order,
            request.sender_name,
            cached_row_indexes,
            historical_row_indexes=historical_indexes,
            pending_replace=pending_replace,
            matched_by=_to_string(process_payload.get("matched_by")),
        )
    effective_order_payload = process_payload.get("effective_order") or final_order.to_excel_dict()
    required_fields = ("单号", "销售员", "客户", "货品名称", "数量", "销售金额", "收货人电话", "收货地址")
    missing_fields = [
        field_name
        for field_name in required_fields
        if _normalize_value(effective_order_payload.get(field_name)) is None
    ]
    needs_review = bool(missing_fields)
    if process_payload.get("success") and process_payload.get("action") in {"updated", "replaced"} and not _has_item_details(final_order):
        missing_fields = [field_name for field_name in missing_fields if field_name not in {"货品名称", "数量", "销售金额"}]
        needs_review = bool(missing_fields)
    product_needs_review = bool(process_payload.get("product_needs_review")) or (
        _to_string(final_order.extra_fields.get("商品名称需复核")) == "是"
    )
    if product_needs_review:
        needs_review = True
        if "货品名称标准化" not in missing_fields:
            missing_fields.append("货品名称标准化")
    result = _json_result(
        success=True,
        action=pipeline_action,
        duplicate_order_number=duplicate_order_number,
        needs_review=needs_review,
        missing_fields=missing_fields,
        message_intent=_to_string(final_order.extra_fields.get("消息意图")),
        parsed_order=parsed.order.to_excel_dict(),
        final_order=final_order.to_excel_dict(),
        process_result=process_payload,
    )
    _append_tool_log(
        "ingest_order_message",
        order_number=final_order.单号,
        result_type=process_payload.get("action"),
        dry_run=request.dry_run,
    )
    return result


@mcp.tool()
def process_excel_order(
    order: ExcelOrder,
    auto_add_new_columns: bool = False,
    dry_run: bool = False,
) -> str:
    """
    将标准订单对象写入 OneDrive 在线 Excel。

    调用规则：
    - 仅接受与你表头一致的订单字段
    - 优先按 `单号` 查找历史记录并更新
    - `单号` 未命中时，按 `客户` 倒序匹配最近一条
    - 仍未命中时，再按 `匹配客户别名` 倒序匹配最近一条
    - 会优先匹配同一个 `销售员` 的最近记录
    - 空值不会覆盖旧值
    - 默认不自动建列，避免误加表头
    - `dry_run=true` 时只做读取和匹配，不真正写入 Excel
    """
    has_item_details = _has_item_details(order)
    message_intent = _to_string(order.extra_fields.get("消息意图")) or ("revise_items" if has_item_details else "supplement")
    replace_existing_block = message_intent == "replace_order"
    try:
        token = get_token_automatically()
    except Exception as exc:
        result = _json_result(
            success=False,
            action="auth_failed",
            error_type="auth_failed",
            message="无法初始化 Microsoft 登录状态，请检查网络或稍后重试。",
            dry_run=dry_run,
            detail=str(exc),
        )
        _append_tool_log("process_excel_order", order_number=order.单号, result_type="auth_failed")
        return result
    if not token:
        result = _json_result(
            success=False,
            action="auth_failed",
            error_type="auth_failed",
            message="无法获取微软授权 Token，请确认已设置环境变量 OC_OD_CLIENT_ID 并完成设备登录。",
            dry_run=dry_run,
        )
        _append_tool_log("process_excel_order", order_number=order.单号, result_type="auth_failed")
        return result

    base_url = _build_base_url()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    proxies = {"http": None, "https": None}
    product_resolution = _standardize_order_products(order, token=token)
    row_dicts = _build_excel_row_dicts(order)
    order_data = row_dicts[0]
    detail_row_count = len(row_dicts)
    has_item_details = _has_item_details(order)

    try:
        resp_cols = requests.get(f"{base_url}/columns", headers=headers, proxies=proxies, timeout=30)
        if resp_cols.status_code != 200:
            error_type, error_message = _classify_graph_error(
                resp_cols.status_code,
                resp_cols.text,
                operation="columns",
            )
            result = _json_result(
                success=False,
                action="read_columns_failed",
                error_type=error_type,
                message=error_message,
                status_code=resp_cols.status_code,
                detail=resp_cols.text,
            )
            _append_tool_log("process_excel_order", order_number=order.单号, result_type=error_type)
            return result
        existing_columns = [col["name"] for col in resp_cols.json().get("value", [])]

        added_columns: list[str] = []
        skipped_columns: list[str] = []
        for key in order_data:
            if key in existing_columns:
                continue
            if auto_add_new_columns:
                add_resp = requests.post(
                    f"{base_url}/columns",
                    headers=headers,
                    json={"name": key},
                    proxies=proxies,
                    timeout=30,
                )
                if add_resp.status_code in (200, 201):
                    existing_columns.append(key)
                    added_columns.append(key)
                else:
                    skipped_columns.append(key)
            else:
                skipped_columns.append(key)

        resp_rows = requests.get(f"{base_url}/rows", headers=headers, proxies=proxies, timeout=30)
        if resp_rows.status_code != 200:
            error_type, error_message = _classify_graph_error(
                resp_rows.status_code,
                resp_rows.text,
                operation="rows",
            )
            result = _json_result(
                success=False,
                action="read_rows_failed",
                error_type=error_type,
                message=error_message,
                status_code=resp_rows.status_code,
                detail=resp_rows.text,
            )
            _append_tool_log("process_excel_order", order_number=order.单号, result_type=error_type)
            return result
        rows_data = resp_rows.json().get("value", [])

        target_order_id = _to_string(order_data.get("单号"))
        target_customer_alias = _to_string(order.匹配客户别名)
        target_customer = _to_string(order_data.get("客户"))
        normalized_target_order_id = _normalize_match_text(target_order_id)
        normalized_target_customer_alias = _normalize_match_text(target_customer_alias)
        normalized_target_customer = _normalize_match_text(target_customer)
        found_row_index = -1
        existing_row_values: list[Any] = []
        matched_by = None
        matched_row_indexes: list[int] = []

        id_col_idx = existing_columns.index("单号") if "单号" in existing_columns else -1
        cust_col_idx = existing_columns.index("客户") if "客户" in existing_columns else -1
        salesperson_col_idx = existing_columns.index("销售员") if "销售员" in existing_columns else -1

        def _matches_target_row(row_vals: list[Any], require_same_salesperson: bool) -> str | None:
            row_id = str(row_vals[id_col_idx]).strip() if 0 <= id_col_idx < len(row_vals) else ""
            row_cust = str(row_vals[cust_col_idx]).strip() if 0 <= cust_col_idx < len(row_vals) else ""
            row_salesperson = str(row_vals[salesperson_col_idx]).strip() if 0 <= salesperson_col_idx < len(row_vals) else ""
            normalized_row_id = _normalize_match_text(row_id)
            normalized_row_cust = _normalize_match_text(row_cust)

            salesperson_ok = True
            if require_same_salesperson and target_salesperson:
                salesperson_ok = _normalize_match_text(row_salesperson) == _normalize_match_text(target_salesperson)

            if normalized_target_order_id and normalized_row_id == normalized_target_order_id:
                return "单号"

            if not salesperson_ok:
                return None

            if normalized_target_customer and normalized_row_cust == normalized_target_customer:
                return "客户"

            if normalized_target_customer_alias and normalized_row_cust == normalized_target_customer_alias:
                return "匹配客户别名"

            return None

        def _collect_matching_row_indexes(start_index: int, require_same_salesperson: bool) -> list[int]:
            matched_indexes = [start_index]

            for i in range(start_index - 1, -1, -1):
                row_vals = rows_data[i].get("values", [[]])[0]
                if _matches_target_row(row_vals, require_same_salesperson):
                    matched_indexes.insert(0, i)
                else:
                    break

            for i in range(start_index + 1, len(rows_data)):
                row_vals = rows_data[i].get("values", [[]])[0]
                if _matches_target_row(row_vals, require_same_salesperson):
                    matched_indexes.append(i)
                else:
                    break

            return matched_indexes

        def _find_matching_row(require_same_salesperson: bool) -> tuple[int, list[Any], str | None, list[int]]:
            for i in range(len(rows_data) - 1, -1, -1):
                row_vals = rows_data[i].get("values", [[]])[0]
                matched_key = _matches_target_row(row_vals, require_same_salesperson)
                if matched_key:
                    matched_indexes = _collect_matching_row_indexes(i, require_same_salesperson)
                    return i, row_vals, matched_key, matched_indexes

            return -1, [], None, []

        target_salesperson = _to_string(order_data.get("销售员"))
        preferred_row_hint = None
        historical_row_hint = order.extra_fields.get("历史订单行索引")
        if isinstance(historical_row_hint, str):
            try:
                historical_row_hint = json.loads(historical_row_hint)
            except Exception:
                historical_row_hint = None
        historical_row_indexes = _normalize_row_indexes(historical_row_hint)
        if replace_existing_block and historical_row_indexes:
            preferred_row_hint = historical_row_indexes
        else:
            preferred_row_hint = order.extra_fields.get("最近草稿行索引")
        if isinstance(preferred_row_hint, str):
            try:
                preferred_row_hint = json.loads(preferred_row_hint)
            except Exception:
                preferred_row_hint = None
        preferred_row_indexes = _normalize_row_indexes(preferred_row_hint)
        if preferred_row_indexes:
            valid_indexes = [idx for idx in preferred_row_indexes if 0 <= idx < len(rows_data)]
            if valid_indexes:
                expanded_indexes = _collect_matching_row_indexes(
                    valid_indexes[-1],
                    require_same_salesperson=bool(target_salesperson),
                )
                if len(expanded_indexes) <= 1 and target_salesperson:
                    expanded_indexes = _collect_matching_row_indexes(
                        valid_indexes[-1],
                        require_same_salesperson=False,
                    )
                matched_row_indexes = expanded_indexes or valid_indexes
                found_row_index = matched_row_indexes[-1]
                existing_row_values = rows_data[found_row_index].get("values", [[]])[0]
                matched_by = "历史订单锚点" if replace_existing_block and historical_row_indexes else "最近草稿"

        if found_row_index == -1:
            found_row_index, existing_row_values, matched_by, matched_row_indexes = _find_matching_row(
                require_same_salesperson=True
            )
        if found_row_index == -1:
            found_row_index, existing_row_values, matched_by, matched_row_indexes = _find_matching_row(
                require_same_salesperson=False
            )

        effective_order = order
        effective_row_dicts = row_dicts
        effective_order_data = order_data
        if matched_row_indexes:
            historical_order = _build_order_from_matched_rows(
                existing_columns=existing_columns,
                rows_data=rows_data,
                matched_row_indexes=matched_row_indexes,
            )
            effective_order = _merge_orders(
                existing_order=historical_order,
                new_order=order,
                sender_name=_to_string(order_data.get("销售员")) or historical_order.销售员,
            )
            effective_row_dicts = _build_excel_row_dicts(effective_order)
            effective_order_data = effective_row_dicts[0]
            order = effective_order
            row_dicts = effective_row_dicts
            order_data = effective_order_data
            detail_row_count = len(row_dicts)
            has_item_details = _has_item_details(order)

        if dry_run:
            preview = [{header: row.get(header, "") for header in EXCEL_HEADERS} for row in row_dicts]
            dry_run_action = "updated" if found_row_index != -1 else "created"
            result = _json_result(
                success=True,
                action="dry_run",
                message="Dry run 已完成，仅解析和匹配，不会写入 Excel。",
                dry_run=True,
                would_action=dry_run_action,
                matched_by=matched_by,
                matched_value=target_order_id or target_customer,
                row_index=found_row_index if found_row_index != -1 else None,
                row_indexes=matched_row_indexes or None,
                historical_row_indexes=historical_row_indexes or None,
                detail_row_count=detail_row_count,
                product_resolution=product_resolution,
                product_needs_review=product_resolution.get("needs_review"),
                order=order_data,
                orders=row_dicts,
                effective_order=order_data,
                excel_row_preview=preview,
                auto_add_new_columns=auto_add_new_columns,
                added_columns=added_columns,
                skipped_columns=skipped_columns,
            )
            _append_tool_log("process_excel_order", order_number=order.单号, result_type="dry_run")
            return result

        if found_row_index != -1:
            if not matched_row_indexes:
                matched_row_indexes = [found_row_index]

            if not has_item_details:
                shared_fields = {
                    "备注",
                    "发货厂家",
                    "产品供应商",
                    "日期",
                    "单号",
                    "销售员",
                    "客户",
                    "运费",
                    "利润",
                    "收货联系人",
                    "收货人电话",
                    "收货地址",
                }
                first_row_index = matched_row_indexes[0]
                total_received = _normalize_value(order.已收)
                total_unpaid = _normalize_value(order.未收)

                for row_index in matched_row_indexes:
                    existing_values = rows_data[row_index].get("values", [[]])[0]
                    new_values: list[Any] = []
                    for idx, col_name in enumerate(existing_columns):
                        old_val = existing_values[idx] if idx < len(existing_values) else ""
                        if col_name in shared_fields:
                            new_val = order_data.get(col_name)
                            new_values.append(new_val if new_val is not None else old_val)
                            continue

                        if col_name == "已收":
                            if row_index == first_row_index:
                                new_values.append(total_received if total_received is not None else old_val)
                            else:
                                new_values.append("")
                            continue

                        if col_name == "未收":
                            if row_index == first_row_index:
                                new_values.append(total_unpaid if total_unpaid is not None else old_val)
                            else:
                                new_values.append("")
                            continue

                        new_values.append(old_val)

                    patch_url = f"{base_url}/rows/itemAt(index={row_index})"
                    resp_patch = requests.patch(
                        patch_url,
                        headers=headers,
                        json={"values": [new_values]},
                        proxies=proxies,
                        timeout=30,
                    )
                    if resp_patch.status_code != 200:
                        error_type, error_message = _classify_graph_error(
                            resp_patch.status_code,
                            resp_patch.text,
                            operation="update",
                        )
                        result = _json_result(
                            success=False,
                            action="update_failed",
                            error_type=error_type,
                            message=error_message,
                            matched_by=matched_by,
                            matched_value=target_order_id or target_customer,
                            row_index=row_index,
                            row_indexes=matched_row_indexes,
                            status_code=resp_patch.status_code,
                            detail=resp_patch.text,
                        )
                        _append_tool_log("process_excel_order", order_number=order.单号, result_type=error_type)
                        return result

                result = _json_result(
                    success=True,
                    action="updated",
                    message="已匹配历史订单并完成补充信息更新。",
                    matched_by=matched_by,
                    matched_value=target_order_id or target_customer,
                    row_index=found_row_index,
                    row_indexes=matched_row_indexes,
                    detail_row_count=len(matched_row_indexes),
                    product_resolution=product_resolution,
                    product_needs_review=product_resolution.get("needs_review"),
                    added_columns=added_columns,
                    skipped_columns=skipped_columns,
                    order=order_data,
                    orders=row_dicts,
                    effective_order=order_data,
                )
                format_ok, format_warning = _format_order_rows(
                    base_url=base_url,
                    headers=headers,
                    proxies=proxies,
                    existing_columns=existing_columns,
                    row_indexes=matched_row_indexes,
                )
                if not format_ok and format_warning:
                    result = _json_result(
                        **json.loads(result),
                        format_warning=format_warning,
                    )
                _append_tool_log("process_excel_order", order_number=order.单号, result_type="updated")
                return result

            if replace_existing_block and matched_row_indexes:
                delete_ok, delete_error_payload = _delete_table_rows(
                    base_url=base_url,
                    headers=headers,
                    proxies=proxies,
                    row_indexes=matched_row_indexes,
                )
                if not delete_ok:
                    _append_tool_log("process_excel_order", order_number=order.单号, result_type="delete_failed")
                    return delete_error_payload or _json_result(
                        success=False,
                        action="delete_failed",
                        error_type="delete_failed",
                        message="删除历史订单块失败。",
                        row_indexes=matched_row_indexes,
                    )

                row_values_list = [[row_data.get(col_name, "") for col_name in existing_columns] for row_data in row_dicts]
                remaining_row_count = len(rows_data) - len(matched_row_indexes)
                created_row_indexes = list(range(remaining_row_count, remaining_row_count + len(row_dicts)))
                resp_post = requests.post(
                    f"{base_url}/rows",
                    headers=headers,
                    json={"values": row_values_list},
                    proxies=proxies,
                    timeout=30,
                )
                if resp_post.status_code == 201:
                    result = _json_result(
                        success=True,
                        action="replaced",
                        message="已删除历史订单块，并按最新明细重建订单。",
                        matched_by=matched_by,
                        matched_value=target_order_id or target_customer,
                        row_index=found_row_index,
                        row_indexes=created_row_indexes,
                        replaced_row_indexes=matched_row_indexes,
                        historical_row_indexes=historical_row_indexes or matched_row_indexes,
                        detail_row_count=detail_row_count,
                        product_resolution=product_resolution,
                        product_needs_review=product_resolution.get("needs_review"),
                        added_columns=added_columns,
                        skipped_columns=skipped_columns,
                        order=order_data,
                        orders=row_dicts,
                        effective_order=order_data,
                        message_intent=message_intent,
                    )
                    format_ok, format_warning = _format_order_rows(
                        base_url=base_url,
                        headers=headers,
                        proxies=proxies,
                        existing_columns=existing_columns,
                        row_indexes=created_row_indexes,
                    )
                    if not format_ok and format_warning:
                        result = _json_result(
                            **json.loads(result),
                            format_warning=format_warning,
                        )
                    _append_tool_log(
                        "process_excel_order",
                        order_number=order.单号,
                        result_type="replaced",
                        matched_by=matched_by,
                        row_indexes=created_row_indexes,
                        replaced_row_indexes=matched_row_indexes,
                        historical_row_indexes=historical_row_indexes or matched_row_indexes,
                        message_intent=message_intent,
                    )
                    return result

                error_type, error_message = _classify_graph_error(
                    resp_post.status_code,
                    resp_post.text,
                    operation="create",
                )
                result = _json_result(
                    success=False,
                    action="create_failed",
                    error_type=error_type,
                    message=error_message,
                    matched_by=matched_by,
                    matched_value=target_order_id or target_customer,
                    row_index=found_row_index,
                    row_indexes=matched_row_indexes,
                    status_code=resp_post.status_code,
                    detail=resp_post.text,
                )
                _append_tool_log("process_excel_order", order_number=order.单号, result_type=error_type)
                return result

            if len(matched_row_indexes) not in (0, detail_row_count):
                result = _json_result(
                    success=False,
                    action="needs_review",
                    error_type="row_count_mismatch",
                    message="匹配到了历史订单，但现有明细行数与本次识别结果不一致，请人工复核后再更新。",
                    matched_by=matched_by,
                    matched_value=target_order_id or target_customer,
                    row_index=found_row_index,
                    row_indexes=matched_row_indexes,
                    historical_row_indexes=historical_row_indexes or matched_row_indexes,
                    detail_row_count=detail_row_count,
                    product_resolution=product_resolution,
                    product_needs_review=product_resolution.get("needs_review"),
                    needs_review=True,
                )
                _append_tool_log(
                    "process_excel_order",
                    order_number=order.单号,
                    result_type="needs_review",
                    matched_by=matched_by,
                    row_indexes=matched_row_indexes,
                    historical_row_indexes=historical_row_indexes or matched_row_indexes,
                    message_intent=message_intent,
                )
                return result

            for row_offset, row_index in enumerate(matched_row_indexes):
                existing_values = rows_data[row_index].get("values", [[]])[0]
                row_data = row_dicts[row_offset]
                new_values: list[Any] = []
                for idx, col_name in enumerate(existing_columns):
                    old_val = existing_values[idx] if idx < len(existing_values) else ""
                    new_val = row_data.get(col_name)
                    new_values.append(new_val if new_val is not None else old_val)

                patch_url = f"{base_url}/rows/itemAt(index={row_index})"
                resp_patch = requests.patch(
                    patch_url,
                    headers=headers,
                    json={"values": [new_values]},
                    proxies=proxies,
                    timeout=30,
                )
                if resp_patch.status_code != 200:
                    error_type, error_message = _classify_graph_error(
                        resp_patch.status_code,
                        resp_patch.text,
                        operation="update",
                    )
                    result = _json_result(
                        success=False,
                        action="update_failed",
                        error_type=error_type,
                        message=error_message,
                        matched_by=matched_by,
                        matched_value=target_order_id or target_customer,
                        row_index=row_index,
                        row_indexes=matched_row_indexes,
                        status_code=resp_patch.status_code,
                        detail=resp_patch.text,
                    )
                    _append_tool_log("process_excel_order", order_number=order.单号, result_type=error_type)
                    return result

            result = _json_result(
                success=True,
                action="updated",
                message="已匹配历史订单并完成多行明细更新。",
                matched_by=matched_by,
                matched_value=target_order_id or target_customer,
                row_index=found_row_index,
                row_indexes=matched_row_indexes,
                detail_row_count=detail_row_count,
                product_resolution=product_resolution,
                product_needs_review=product_resolution.get("needs_review"),
                added_columns=added_columns,
                skipped_columns=skipped_columns,
                order=order_data,
                orders=row_dicts,
                effective_order=order_data,
            )
            format_ok, format_warning = _format_order_rows(
                base_url=base_url,
                headers=headers,
                proxies=proxies,
                existing_columns=existing_columns,
                row_indexes=matched_row_indexes,
            )
            if not format_ok and format_warning:
                result = _json_result(
                    **json.loads(result),
                    format_warning=format_warning,
                )
            _append_tool_log("process_excel_order", order_number=order.单号, result_type="updated")
            return result

        row_values_list = [[row_data.get(col_name, "") for col_name in existing_columns] for row_data in row_dicts]
        created_row_indexes = list(range(len(rows_data), len(rows_data) + len(row_dicts)))
        resp_post = requests.post(
            f"{base_url}/rows",
            headers=headers,
            json={"values": row_values_list},
            proxies=proxies,
            timeout=30,
        )
        if resp_post.status_code == 201:
            result = _json_result(
                success=True,
                action="created",
                message="未匹配到历史记录，已按商品明细新增订单。",
                matched_by=None,
                matched_value=None,
                row_index=None,
                detail_row_count=detail_row_count,
                product_resolution=product_resolution,
                product_needs_review=product_resolution.get("needs_review"),
                added_columns=added_columns,
                skipped_columns=skipped_columns,
                order=order_data,
                orders=row_dicts,
                effective_order=order_data,
            )
            format_ok, format_warning = _format_order_rows(
                base_url=base_url,
                headers=headers,
                proxies=proxies,
                existing_columns=existing_columns,
                row_indexes=created_row_indexes,
            )
            if not format_ok and format_warning:
                result = _json_result(
                    **json.loads(result),
                    format_warning=format_warning,
                )
            _append_tool_log("process_excel_order", order_number=order.单号, result_type="created")
            return result
        error_type, error_message = _classify_graph_error(
            resp_post.status_code,
            resp_post.text,
            operation="create",
        )
        result = _json_result(
            success=False,
            action="create_failed",
            error_type=error_type,
            message=error_message,
            status_code=resp_post.status_code,
            detail=resp_post.text,
            added_columns=added_columns,
            skipped_columns=skipped_columns,
        )
        _append_tool_log("process_excel_order", order_number=order.单号, result_type=error_type)
        return result

    except Exception as exc:
        result = _json_result(
            success=False,
            action="exception",
            message="脚本运行异常。",
            detail=str(exc),
        )
        _append_tool_log("process_excel_order", order_number=order.单号, result_type="exception")
        return result


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Chengyi order MCP server for WeChat order parsing and Excel writing.",
    )
    parser.add_argument(
        "--host",
        default=DEFAULT_HOST,
        help=f"Host to bind to (default: {DEFAULT_HOST})",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=DEFAULT_PORT,
        help=f"Port to listen on (default: {DEFAULT_PORT})",
    )
    parser.add_argument(
        "--transport",
        choices=["streamable-http", "stdio"],
        default=os.getenv("CY_EXCEL_MCP_TRANSPORT", "streamable-http"),
        help="Transport type (default: streamable-http)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        if args.transport == "streamable-http":
            print(f"Chengyi_Order_Manager MCP listening on http://{args.host}:{args.port}/mcp")
            mcp.run(transport="streamable-http")
        else:
            mcp.run(transport="stdio")
    except KeyboardInterrupt:
        # Suppress the asyncio/anyio stack trace on manual Ctrl+C shutdown.
        print("Chengyi_Order_Manager MCP stopped")


if __name__ == "__main__":
    main()
