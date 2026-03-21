import atexit
import argparse
import json
import os
import re
from datetime import datetime
from typing import Any

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
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_HOST = os.getenv("CY_EXCEL_MCP_HOST", "127.0.0.1")
DEFAULT_PORT = int(os.getenv("CY_EXCEL_MCP_PORT", "18061"))
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


def _build_msal_http_client() -> requests.Session:
    session = requests.Session()
    # Avoid inheriting broken system proxy settings during Microsoft login.
    session.trust_env = False
    return session


def _json_result(**payload: Any) -> str:
    return json.dumps(payload, ensure_ascii=False)


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


def _extract_first(pattern: str, text: str, flags: int = 0) -> str | None:
    match = re.search(pattern, text, flags)
    if not match:
        return None
    return match.group(1).strip()


def _extract_phone(text: str) -> str | None:
    match = re.search(r"(?<!\d)(1\d{10})(?!\d)", text)
    return match.group(1) if match else None


def _same_text(left: str | None, right: str | None) -> bool:
    return _to_string(left) == _to_string(right)


def _extract_order_number_and_customer_alias(text: str) -> tuple[str | None, str | None]:
    raw_order_number = _extract_first(r"单号[:：]\s*(.+)", text)
    if raw_order_number is None:
        return None, None

    normalized_order_number = raw_order_number.strip()
    alias = None
    alias_match = re.search(r"[（(]([^()（）]+)[)）]\s*$", normalized_order_number)
    if alias_match:
        alias = _normalize_value(alias_match.group(1))
        normalized_order_number = re.sub(r"\s*[（(][^()（）]+[)）]\s*$", "", normalized_order_number).strip()

    return _normalize_value(normalized_order_number), alias


def _extract_product_line(text: str) -> tuple[str | None, str | None, str | None, str | None]:
    for line in [item.strip() for item in text.splitlines() if item.strip()]:
        if any(keyword in line for keyword in ("单号", "收件人", "手机", "电话", "地址")):
            continue
        match = re.search(
            r"^(?P<name>.+?)(?P<qty>\d+(?:\.\d+)?)\s*(?P<unit>[^\d\s元]+)?\s*(?P<amount>\d+(?:\.\d+)?)\s*元?$",
            line,
        )
        if match:
            return (
                match.group("name").strip(),
                match.group("qty").strip(),
                _normalize_value(match.group("unit")),
                match.group("amount").strip(),
            )
    return None, None, None, None


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


def get_token_automatically() -> str | None:
    client_id = os.getenv("OC_OD_CLIENT_ID")
    if not client_id:
        return None

    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(_read_text_file(CACHE_FILE))

    def _persist_cache() -> None:
        if cache.has_state_changed:
            _write_text_file(CACHE_FILE, cache.serialize())

    atexit.register(_persist_cache)

    app = msal.PublicClientApplication(
        client_id,
        authority=AUTHORITY,
        token_cache=cache,
        http_client=_build_msal_http_client(),
    )

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
    return f"{GRAPH_ROOT}/me/drive/root:/{FILE_PATH}:/workbook/tables('{TABLE_NAME}')"


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
    extra_fields: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="after")
    def validate_match_keys(self) -> "ExcelOrder":
        if not _normalize_value(self.单号) and not _normalize_value(self.客户):
            raise ValueError("至少提供 `单号` 或 `客户`。")
        return self

    def finalize(self) -> "ExcelOrder":
        total_payment = _to_float(self.总货款)
        received_payment = _to_float(self.已收)
        sales_amount = _to_float(self.销售金额)
        shipping_fee = _to_float(self.运费)
        cost_amount = _to_float(self.成本金额)

        if total_payment is None and sales_amount is not None:
            total_payment = sales_amount
        if received_payment is None and sales_amount is not None:
            received_payment = sales_amount
        unpaid = None
        if total_payment is not None and received_payment is not None:
            unpaid = total_payment - received_payment

        if sales_amount is None and total_payment is not None:
            sales_amount = total_payment

        profit = _to_float(self.利润)
        if profit is None and sales_amount is not None:
            cost_value = cost_amount or 0.0
            shipping_value = shipping_fee or 0.0
            profit = sales_amount - cost_value - shipping_value

        self.日期 = _normalize_date(self.日期, None)
        self.销售金额 = _format_money(sales_amount)
        self.总货款 = _format_money(total_payment)
        self.已收 = _format_money(received_payment)
        self.未收 = _format_money(unpaid)
        self.利润 = _format_money(profit)
        self.销售单价 = _format_money(_to_float(self.销售单价))
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


def _parse_wechat_order_message_model(
    raw_message: str,
    sender_name: str | None = None,
    group_name: str | None = None,
    message_time: str | None = None,
) -> ParsedWechatOrder:
    lines = [line.strip() for line in raw_message.splitlines() if line.strip()]

    order_number, customer_alias = _extract_order_number_and_customer_alias(raw_message)
    contact_name = _extract_first(r"收件人[:：]\s*(.+)", raw_message)
    phone = _extract_phone(raw_message)
    address = _extract_first(r"收货地址[:：]\s*(.+)", raw_message)
    order_date = _normalize_date(order_number, message_time)
    sales_amount = _extract_first(r"(?:全款|总货款)[:：]?\s*([0-9]+(?:\.[0-9]+)?)\s*元?", raw_message)
    payment_note = _extract_first(r"(.+收款)", raw_message)

    product_name, quantity, quantity_unit, product_line_amount = _extract_product_line(raw_message)
    if sales_amount is None:
        sales_amount = product_line_amount

    customer = None
    for line in lines:
        if line.startswith(("单号", "收件人", "手机号码", "手机号", "电话", "收货地址", "全款", "总货款")):
            continue
        if payment_note and line == payment_note:
            continue
        if line == product_name:
            continue
        if re.search(r"\d", line) and product_name and line.find(product_name) != -1:
            continue
        customer = line
        break

    if customer is None:
        customer = customer_alias or contact_name

    order = ExcelOrder(
        日期=order_date,
        单号=order_number,
        匹配客户别名=_normalize_value(customer_alias),
        销售员=_normalize_value(sender_name),
        客户=_normalize_value(customer),
        货品名称=_normalize_value(product_name),
        数量=_normalize_value(quantity),
        数量单位=_normalize_value(quantity_unit),
        销售金额=_normalize_value(sales_amount),
        总货款=_normalize_value(sales_amount),
        已收=_normalize_value(sales_amount),
        收货联系人=_normalize_value(contact_name or customer),
        收货人电话=_normalize_value(phone),
        收货地址=_normalize_value(address),
        备注=_normalize_value(payment_note),
    )

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

    merged_extra = dict(existing_order.extra_fields)
    for key, value in new_order.extra_fields.items():
        normalized_key = str(key).strip()
        if normalized_key:
            merged_extra[normalized_key] = _merge_prefer_new(merged_extra.get(normalized_key), value)
    merged.extra_fields = merged_extra
    return merged


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
    return parsed.model_dump_json(indent=2, exclude_none=True, ensure_ascii=False)


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
    return _json_result(
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

    if request.existing_order is not None:
        final_order = _merge_orders(
            existing_order=request.existing_order,
            new_order=parsed.order,
            sender_name=request.sender_name,
        )
        pipeline_action = "merged_then_processed"
        duplicate_order_number = bool(
            _normalize_value(request.existing_order.单号)
            and _normalize_value(parsed.order.单号)
            and _same_text(request.existing_order.单号, parsed.order.单号)
        )
    else:
        final_order = parsed.order
        pipeline_action = "parsed_then_processed"
        duplicate_order_number = False

    missing_fields = [
        field_name
        for field_name in ("单号", "销售员", "客户", "货品名称", "数量", "销售金额", "收货人电话", "收货地址")
        if _normalize_value(getattr(final_order, field_name)) is None
    ]
    needs_review = bool(missing_fields)
    process_result = process_excel_order(
        order=final_order,
        auto_add_new_columns=request.auto_add_new_columns,
    )

    return _json_result(
        success=True,
        action=pipeline_action,
        duplicate_order_number=duplicate_order_number,
        needs_review=needs_review,
        missing_fields=missing_fields,
        parsed_order=parsed.order.to_excel_dict(),
        final_order=final_order.to_excel_dict(),
        process_result=json.loads(process_result),
    )


@mcp.tool()
def process_excel_order(
    order: ExcelOrder,
    auto_add_new_columns: bool = False,
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
    """
    token = get_token_automatically()
    if not token:
        return _json_result(
            success=False,
            action="auth_failed",
            message="无法获取微软授权 Token，请确认已设置环境变量 OC_OD_CLIENT_ID 并完成设备登录。",
        )

    order_data = order.to_excel_dict()
    base_url = _build_base_url()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    proxies = {"http": None, "https": None}

    try:
        resp_cols = requests.get(f"{base_url}/columns", headers=headers, proxies=proxies, timeout=30)
        if resp_cols.status_code != 200:
            return _json_result(
                success=False,
                action="read_columns_failed",
                message="无法读取 Excel 表头。",
                status_code=resp_cols.status_code,
                detail=resp_cols.text,
            )
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
            return _json_result(
                success=False,
                action="read_rows_failed",
                message="无法读取 Excel 行数据。",
                status_code=resp_rows.status_code,
                detail=resp_rows.text,
            )
        rows_data = resp_rows.json().get("value", [])

        target_order_id = _to_string(order_data.get("单号"))
        target_customer_alias = _to_string(order.匹配客户别名)
        target_customer = _to_string(order_data.get("客户"))
        found_row_index = -1
        existing_row_values: list[Any] = []
        matched_by = None

        id_col_idx = existing_columns.index("单号") if "单号" in existing_columns else -1
        cust_col_idx = existing_columns.index("客户") if "客户" in existing_columns else -1
        salesperson_col_idx = existing_columns.index("销售员") if "销售员" in existing_columns else -1

        def _find_matching_row(require_same_salesperson: bool) -> tuple[int, list[Any], str | None]:
            for i in range(len(rows_data) - 1, -1, -1):
                row_vals = rows_data[i].get("values", [[]])[0]
                row_id = str(row_vals[id_col_idx]).strip() if 0 <= id_col_idx < len(row_vals) else ""
                row_cust = str(row_vals[cust_col_idx]).strip() if 0 <= cust_col_idx < len(row_vals) else ""
                row_salesperson = str(row_vals[salesperson_col_idx]).strip() if 0 <= salesperson_col_idx < len(row_vals) else ""

                salesperson_ok = True
                if require_same_salesperson and target_salesperson:
                    salesperson_ok = row_salesperson == target_salesperson

                if target_order_id and row_id == target_order_id:
                    return i, row_vals, "单号"

                if not salesperson_ok:
                    continue

                if target_customer and row_cust == target_customer:
                    return i, row_vals, "客户"

                if target_customer_alias and row_cust == target_customer_alias:
                    return i, row_vals, "匹配客户别名"

            return -1, [], None

        target_salesperson = _to_string(order_data.get("销售员"))
        found_row_index, existing_row_values, matched_by = _find_matching_row(require_same_salesperson=True)
        if found_row_index == -1:
            found_row_index, existing_row_values, matched_by = _find_matching_row(require_same_salesperson=False)

        if found_row_index != -1:
            new_values: list[Any] = []
            for idx, col_name in enumerate(existing_columns):
                old_val = existing_row_values[idx] if idx < len(existing_row_values) else ""
                new_val = order_data.get(col_name)
                new_values.append(new_val if new_val is not None else old_val)

            patch_url = f"{base_url}/rows/itemAt(index={found_row_index})"
            resp_patch = requests.patch(
                patch_url,
                headers=headers,
                json={"values": [new_values]},
                proxies=proxies,
                timeout=30,
            )
            if resp_patch.status_code == 200:
                return _json_result(
                    success=True,
                    action="updated",
                    message="已匹配历史订单并完成无损更新。",
                    matched_by=matched_by,
                    matched_value=target_order_id or target_customer,
                    row_index=found_row_index,
                    added_columns=added_columns,
                    skipped_columns=skipped_columns,
                    order=order_data,
                )
            return _json_result(
                success=False,
                action="update_failed",
                message="已找到历史记录，但更新失败。",
                matched_by=matched_by,
                matched_value=target_order_id or target_customer,
                row_index=found_row_index,
                status_code=resp_patch.status_code,
                detail=resp_patch.text,
            )

        row_values = [order_data.get(col_name, "") for col_name in existing_columns]
        resp_post = requests.post(
            f"{base_url}/rows",
            headers=headers,
            json={"values": [row_values]},
            proxies=proxies,
            timeout=30,
        )
        if resp_post.status_code == 201:
            return _json_result(
                success=True,
                action="created",
                message="未匹配到历史记录，已新增订单。",
                matched_by=None,
                matched_value=None,
                row_index=None,
                added_columns=added_columns,
                skipped_columns=skipped_columns,
                order=order_data,
            )
        return _json_result(
            success=False,
            action="create_failed",
            message="新增订单失败。",
            status_code=resp_post.status_code,
            detail=resp_post.text,
            added_columns=added_columns,
            skipped_columns=skipped_columns,
        )

    except Exception as exc:
        return _json_result(
            success=False,
            action="exception",
            message="脚本运行异常。",
            detail=str(exc),
        )


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
