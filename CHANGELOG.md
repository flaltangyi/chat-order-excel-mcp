# Changelog

All notable changes to this project will be documented in this file.
本项目的重要变更都会记录在这个文件中。

## [Unreleased]

### Changed / 变更

- Quantity units are now parsed generically from the text immediately following the quantity number, while parenthetical notes are ignored for the unit field
- 数量单位现在统一按“数字后面的文本”通用提取，括号中的补充说明不再写入单位列

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
