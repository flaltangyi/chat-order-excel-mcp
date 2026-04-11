# Changelog

All notable changes to this project will be documented in this file.
本项目的重要变更都会记录在这个文件中。

## [Unreleased]

### Added / 新增

- Added a repository `TODO.md` with high-, medium-, and low-priority roadmap items
- 新增仓库级 `TODO.md`，按高、中、低优先级整理后续路线图
- Added a local recent-draft cache for split image/text order follow-ups so supplement messages can reuse the latest matching draft before Excel matching
- 新增本地“最近草稿缓存”，用于处理图片和补充文字分开发送的场景，让后续补充消息在匹配 Excel 之前先复用最近草稿
- Added `.gitignore` coverage for the runtime draft cache file `order_draft_cache.json`
- 为运行期草稿缓存文件 `order_draft_cache.json` 补充了 `.gitignore` 忽略规则

### Changed / 变更

- Quantity units are now parsed generically from the text immediately following the quantity number, while parenthetical notes are ignored for the unit field
- 数量单位现在统一按“数字后面的文本”通用提取，括号中的补充说明不再写入单位列
- Existing multi-line orders are now updated by collecting the contiguous order block around the matched row, instead of only scanning downward
- 更新已有多行订单时，现在会围绕命中行向上和向下收集连续订单块，而不再只向下扫描
- Standalone address lines can now be recognized without requiring a `收货地址:` prefix
- 现在无需 `收货地址:` 前缀，也可以识别独立地址行
- Supplement-only follow-up messages can now update shared fields across an existing multi-line order block without requiring item-line counts to match
- 仅补充单号/收款/联系人/地址的后续消息，现在可以直接更新已有多行订单块的共享字段，不再要求商品行数一致
- Historical multi-line order blocks are now rebuilt into typed `ExcelOrder` objects using safer string coercion for shared fields such as dates and phone numbers
- 回填历史多行订单块时，现在会把日期、电话等共享字段安全地转换为字符串，再构造成 `ExcelOrder`，避免类型校验中断补单
- Split-send supplement messages now try to inherit and merge the latest draft block for the same salesperson and customer/order before falling back to pure Excel matching
- 图片先发、文字后补的消息现在会优先继承并合并同一销售员下最近的草稿订单块，再回退到 Excel 匹配
- Replace-confirmation messages such as `客户名以这个为准` now normalize the customer target correctly, cache pending replacement drafts, and expand cached row hints back into the full historical order block before replacement
- `客户名以这个为准` 这类确认替换消息现在会正确提取客户名、缓存待替换草稿，并在真正替换前把缓存行索引扩展回完整历史订单块
- Replace-order deletion now uses working Graph row deletion paths, allowing old multi-line blocks to be removed before rebuilding the latest single-line or revised order
- 替换订单时现在改用可用的 Graph 行删除路径，能够先删除旧的多行订单块，再重建最新的一行或修订版订单

## [0.1.2] - 2026-03-30

### Added / 新增

- Public OpenClaw integration docs in English and Chinese
- 新增 OpenClaw 接入文档的中英文说明
- Documented the recommended `orderentry + wecom + stdio MCP` topology
- 补充推荐的 `orderentry + wecom + stdio MCP` 接入拓扑
- Documented the security boundary between shared order-entry bots and admin-only host control
- 补充共享订单机器人与管理员宿主机控制之间的权限边界
- Added a repository pre-commit hook to require staged `CHANGELOG.md` updates
- 新增仓库级 pre-commit hook，要求提交其他改动时必须同步暂存 `CHANGELOG.md`

### Changed / 变更

- Improved multi-item OCR parsing for image-based orders
- 增强图片订单的多商品 OCR 文本解析
- Multi-line item orders are now written as one Excel row per product line
- 多商品订单现在按“每个商品一行”写入 Excel
- Duplicate matching now normalizes order number, customer name, and alias text before comparison
- 重复单匹配现在会先标准化单号、客户名和客户别名后再比较
- Missing dates now prefer parsing from compact order numbers such as `260313-05`, otherwise fall back to the current date
- 日期缺失时会优先从 `260313-05` 这类单号中提取，否则回退为当天日期
- Removed merge-based payment formatting for Excel tables and switched to highlight styling blocks
- 删除 Excel Table 内的合并单元格方案，改为高亮样式块显示已收/未收
- Fixed Microsoft Graph formatting calls by updating alignment through `PATCH /format`
- 修复 Microsoft Graph 样式调用，对齐改为通过 `PATCH /format` 更新

## [0.1.1] - 2026-03-21

### Added / 新增

- Chinese runtime guide at `docs/RUNNING.zh-CN.md`
- 新增中文运行说明 `docs/RUNNING.zh-CN.md`
- Final delivery checklist for OpenClaw deployment and verification
- 新增 OpenClaw 部署和验证的最终交付清单

### Changed / 变更

- Startup script now writes logs into `logs/`
- 启动脚本现在会把日志写入 `logs/`
- Log filenames now include date, time, and per-day sequence number
- 日志文件名现在包含日期、时间和每日序号
- Old log files are deleted automatically based on retention days
- 旧日志会按照保留天数自动清理
- Startup script now attempts to clear stale MCP processes before relaunch
- 启动脚本在重启前会尝试清理陈旧的 MCP 进程
- Runtime documentation now includes validated personal OneDrive guidance
- 运行文档现在包含已验证通过的个人 OneDrive 配置说明

## [0.1.0] - 2026-03-17

### Added / 新增

- Initial MCP server for parsing WeChat order messages
- 初始版本提供用于解析微信群订单消息的 MCP 服务
- Order merge flow for follow-up message updates
- 支持后续补充消息的订单合并流程
- OneDrive Excel write and update support through Microsoft Graph
- 支持通过 Microsoft Graph 写入和更新 OneDrive Excel
- Streamable HTTP support for OpenClaw MCP integration
- 支持用于 OpenClaw 接入的 Streamable HTTP MCP 方式
- Bootstrap script, environment template, and MCP config example
- 提供 bootstrap 脚本、环境变量模板和 MCP 配置示例
- English and Chinese README documents
- 提供中英文 README 文档
