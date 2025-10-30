# README_REFACTOR

此分支包含从你提供的源码快照出发的重构结果（初始批次）。目标是：整理模块、合并重复代码、提供更清晰的架构，并确保源码以 ANSI/GBK 编码保存以便 VB6 直接打开。

主要说明：
- 分支：refactor/snapshot-2025-10-30
- 默认包含 Mock 模式用于离线调试（可在 CONFIG.INI 或代码中开启/关闭）
- ffmpeg 未打包，请手动下载并放置到 tools\ffmpeg\ 或在 CONFIG.INI 中设置 ffmpeg 路径

如何在本机运行：
1. 下载并解压项目
2. 在 PowerShell 中运行 scripts\convert_to_ansi.ps1（可选，文件已为 ANSI）
3. 用 VB6 打开项目文件 (.vbp)
4. 补齐缺失的 OCX/COM 控件（MSHFlexGrid 等）
5. 运行（F5）并在调试输出中观察日志

如遇问题请在仓库 issue 中描述具体错误日志与截图。

