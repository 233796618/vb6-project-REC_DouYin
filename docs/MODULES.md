# MODULES 说明

- bas/GridHelper.bas
  - 功能：统一 MSHFlexGrid 操作，列宽管理，批量/缓冲更新
  - 主要接口：InitializeGrid(grid)、AutoSizeColumns(grid)、UpdateGridWithBuffer(grid, rooms, buffer)

- bas/M__Global.bas
  - 功能：全局单例管理（Scheduler、LiveroomManager）、应用启动与优雅关闭
  - 主要接口：InitializeGlobalVariables、SafeShutdownApplication、CompleteShutdown

- scripts/convert_to_ansi.ps1
  - 功能：将源码文件转换为 ANSI/GBK 编码以便 VB6 打开

- scripts/package_project.ps1
  - 功能：把工程打包为 dist/project_YYYYMMDD_HHMMSS.zip

后续计划：合并 TimerManager、PipeManager、ProcessRunner、TaskScheduler 等核心模块，补充更多中文注释与示例数据。

