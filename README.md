# dsh-mini PowerShell

PowerShell 原生运行时版本的 DSH 极简 Agent。它将原来的 Python 单文件运行时迁移为一个独立的 `dsh-mini.ps1`，不依赖 Python、exe 或第三方 PowerShell 模块。

## 功能

- OpenAI Chat Completions 兼容接口
- SSE 流式输出、重试、模型列表扫描
- `pwsh` 工具：独立 .NET Runspace，变量、函数和当前目录跨调用保持
- `str_replace_editor`：`view`、`create`、`str_replace`、`insert`
- 交互式 CLI、单次 `-Prompt` 执行、WinForms GUI
- 配置、会话保存、离线自检和 PowerShell 诊断
- 支持 Windows PowerShell 5.1 和 PowerShell 7+

## 快速开始

第一次运行配置向导：

```powershell
powershell -ExecutionPolicy Bypass -File .\dsh-mini.ps1 -Setup -Cli
```

启动 GUI：

```powershell
powershell -ExecutionPolicy Bypass -File .\dsh-mini.ps1 -Gui
```

单次执行：

```powershell
pwsh -File .\dsh-mini.ps1 -Prompt '列出当前目录' -Cli
```

## `irm` 一行启动

```powershell
irm https://raw.githubusercontent.com/snake-aabb-wtf/dsh-mini-powershell/main/dsh-mini.ps1 | iex
```

带参数时：

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/snake-aabb-wtf/dsh-mini-powershell/main/dsh-mini.ps1))) -Prompt '列出当前目录' -Cli
```

远程执行前请审阅脚本内容。通过 `irm | iex` 运行时，配置默认保存在当前目录；当前目录不可写时回退到 `%APPDATA%\dsh-mini\`。

## 配置

复制 `dsh-mini.config.ps.example.json` 为 `dsh-mini.config.json`，或直接运行 `-Setup`。也可以使用环境变量覆盖基础配置：

- `DSH_MINI_BASE_URL`
- `DSH_MINI_API_KEY`
- `DSH_MINI_MODEL`

常用参数：`-ShellMode persistent|oneshot`、`-NoStream`、`-NoPicker`、`-SelfTest -Cli`、`-ShellCheck -Cli`。

## 文件

- `dsh-mini.ps1`：完整 PowerShell 运行时
- `dsh-mini.config.ps.example.json`：配置示例
- `dsh-mini-irm.txt`：一行启动示例

本仓库暂不附带许可证文件；代码版权状态未作声明。
