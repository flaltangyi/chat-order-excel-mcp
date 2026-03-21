# Changelog

All notable changes to this project will be documented in this file.

## [0.1.1] - 2026-03-21

### Added

- Chinese runtime guide at `docs/RUNNING.zh-CN.md`
- Final delivery checklist for OpenClaw deployment and verification

### Changed

- Startup script now writes logs into `logs/`
- Log filenames now include date, time, and per-day sequence number
- Old log files are deleted automatically based on retention days
- Startup script now attempts to clear stale MCP processes before relaunch
- Runtime documentation now includes validated personal OneDrive guidance

## [0.1.0] - 2026-03-17

### Added

- Initial MCP server for parsing WeChat order messages
- Order merge flow for follow-up message updates
- OneDrive Excel write and update support through Microsoft Graph
- Streamable HTTP support for OpenClaw MCP integration
- Bootstrap script, environment template, and MCP config example
- English and Chinese README documents
