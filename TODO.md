# TODO

This file tracks the next recommended improvements for `chat-order-excel-mcp`.

本文档用于记录 `chat-order-excel-mcp` 后续建议补强的功能。

## High Priority / 高优先级

- Add a dedicated `parse_order_ocr_text` entrypoint so OCR text and normal chat text can be handled with different parsing rules.
- 增加专门的 `parse_order_ocr_text` 入口，把 OCR 文本和普通聊天文本分开处理。

- Add salesperson and customer-alias mapping tables to improve matching stability across nicknames and abbreviations.
- 增加销售员映射表和客户别名映射表，提升昵称和缩写场景下的匹配稳定性。

- Add an order-block repair tool to detect and fix duplicate order numbers, fragmented multi-line orders, and malformed historical rows.
- 增加订单块修复工具，用于扫描和修复重复单号、断裂的多行订单块以及历史脏数据。

- Extend message-intent classification so the system can distinguish between `supplement`, `revise_items`, and `replace_order`, and handle “以这个为准 / 这个为主” as a strong replacement signal.
- 继续增强消息意图分类，让系统区分 `supplement`、`revise_items`、`replace_order`，并把“以这个为准 / 这个为主”视为强替换信号。

- Improve latest-block replacement rules so when a salesperson sends a new image/text and explicitly says “以这个为准”, the latest matching order block can be safely superseded.
- 增强最新订单块替换规则：当业务员发送新图片/新文本并明确说“以这个为准”时，可以安全地用最新订单块覆盖旧块语义。

- Add a long-term order-change mode that prefers Excel history matching over recent draft cache when a salesperson sends a quoted follow-up such as “客户名以这个为准”, especially when the order may be revised days later and the order number is missing.
- 增加“长周期改单模式”：当业务员引用前面的图片或文本并发送“客户名以这个为准”这类说明时，即使已过多日、且单号缺失，也优先依赖 Excel 历史订单块匹配，而不是只依赖最近草稿缓存。

- In long-term change mode, use `销售员 + 客户 + 匹配客户别名` as the primary replacement key and treat recent-draft cache only as a short-term accelerator.
- 在长周期改单模式里，以 `销售员 + 客户 + 匹配客户别名` 作为主要替换锚点，最近草稿缓存只作为短期加速器，不再作为唯一依据。

- Add automated regression tests for OCR parsing, multi-line updates, duplicate matching, and supplement-only updates.
- 增加自动化回归测试，覆盖 OCR 解析、多行更新、重复单匹配和补充消息更新。

## Medium Priority / 中优先级

- Add a debug mode such as `CY_EXCEL_MCP_DEBUG=1` to expose matching details, overwrite decisions, and formatting fallbacks.
- 增加调试模式，例如 `CY_EXCEL_MCP_DEBUG=1`，输出匹配细节、字段覆盖决策和样式降级信息。

- Expand request logs to include matched row indexes, update strategy, and whether a message was treated as a supplement or a new order.
- 扩展请求日志，记录命中的行索引、更新策略，以及当前消息被判定为补充还是新单。

- Add an explicit “update only the latest draft block” option for OpenClaw workflows.
- 增加“只更新最近草稿块”的显式参数，方便 OpenClaw 精确控制更新行为。

- Add a configurable long-term replacement window and audit trail so delayed order changes can record which historical block was replaced and why.
- 增加可配置的长周期替换窗口和审计记录，让延迟改单时能记录替换了哪一块历史订单、依据是什么。

- Add structured payment metadata so notes such as “微信收款 / 支付宝收款 / 张三收款” can be separated from general remarks.
- 增加结构化收款信息，把“微信收款 / 支付宝收款 / 张三收款”等从普通备注里拆出来。

- Improve Excel formatting rules for warning rows, unpaid rows, and duplicate-order review states.
- 增强 Excel 样式规则，对待复核、未收款、重复单等状态做更清晰的高亮。

## Low Priority / 低优先级

- Add a `Makefile` or unified task commands such as `make run`, `make test`, and `make lint`.
- 增加 `Makefile` 或统一任务命令，例如 `make run`、`make test`、`make lint`。

- Add more desensitized sample requests and expected outputs for documentation and testing.
- 增加更多脱敏示例请求和预期输出，用于文档和测试。

- Add a GitHub release workflow to standardize changelog, tagging, and release-note generation.
- 增加 GitHub Release 流程，规范化 changelog、tag 和 release note 的生成。

- Add startup self-check tools that validate workbook existence, table names, and required headers before live use.
- 增加启动自检工具，在正式使用前验证 workbook、table 名称和关键表头是否存在。

- Consider a separate admin-only utility agent for repair, migration, and historical cleanup tasks.
- 可以考虑单独做一个仅管理员可用的工具型 Agent，用于修复、迁移和历史清理。
