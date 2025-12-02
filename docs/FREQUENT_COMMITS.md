FREQUENT COMMITS LOG
Timestamp (UTC): 2025-12-02T00:00:00Z
Branch: refactor/snapshot-2025-10-30

Objective: 按用户要求开始“频繁提交”策略以恢复信任。每次提交都将是小、可验证的原子变更，并附带验证步骤或 smoke-test。

计划（默认频率）：每 1 小时一个小提交，连续 24 小时，之后根据结果调整频率。

本次提交（起始提交）：
- 文件：docs/FREQUENT_COMMITS.md（此文件）
- 说明：记录频繁提交策略的启动与可检验计划。

后续预期原子提交（示例）：
1) pipe/realmode: util/PipeManager 的最小可测真实模式实现（CreateProcess + stdout/stderr 重定向 + 非阻塞读取 + 资源清理），含 smoke_test 脚本。
2) process/runner: util/ProcessRunner 的最小可测实现与示例。
3) task/skeleton: core/TaskQueue, core/TaskManager 的骨架实现与单元 smoke-test。
4) ui/mock-integration: ui/frmMain 的 Mock 集成入口（加载 sample_rooms.json, 启动 Mock 任务流）。
5) docs/tests: README 更新与 smoke_test 脚本说明。

验收与回报机制：
- 每次提交后我会在此会话贴出 commit sha, 变更文件清单与本地验证步骤（如何运行 smoke_test 以及预期输出）。
- 若某次提交无法在 1 小时内完成，我会在 1 小时内说明延迟原因与新的 ETA。

Notes:
- 不包含 ffmpeg 二进制。请将本地 ffmpeg 放到 tools/ffmpeg 并在 README 中设置路径。
- 你已确认：Win10, 已注册 MSHFlexGrid。
