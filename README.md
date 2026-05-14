# Windows 自动日出/日落主题切换

基于北京日出、日落时间自动切换 Windows 浅色/深色主题。日出后使用浅色模式，日落后使用深色模式，并通过计划任务处理每日刷新、定时切换和开机/登录补偿。

## 功能特性

- 按北京坐标（39.9042, 116.4074）从 sunrise-sunset API 获取日出、日落时间。
- 每天自动更新当天和次日的浅色/深色切换任务。
- 开机或登录后自动执行 `Auto` 补偿，修复错过切换时间后的主题状态。
- 即使注册表已经是目标主题，`Auto` 补偿也会强制刷新外观，避免任务栏不同步。
- sunrise-sunset API 失败时优先使用本地缓存，再回退到 7:00 / 18:00。
- 手动 `Light` / `Dark` 切换默认只做轻量刷新，不重启 `explorer.exe`。

## 文件说明

- **AutoThemeSwitcher.ps1**
  - 执行主题切换逻辑。
  - 支持参数：`-Mode Auto|Light|Dark`。
  - 写入注册表：
    - `HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize`
    - `SystemUsesLightTheme` / `AppsUseLightTheme`
  - `Auto` 模式会强制刷新外观，并允许重启 `explorer.exe` 作为任务栏同步兜底。

- **UpdateSchedule.ps1**
  - 获取北京日出、日落时间。
  - 创建或更新两个一次性任务：
    - `\AutoTheme\AutoThemeSwitcher-Light`
    - `\AutoTheme\AutoThemeSwitcher-Dark`
  - 如果当天时间已过，会自动排到次日。
  - API 失败时优先读取 `SunTimesCache.json`，再使用固定兜底时间。

- **InstallTask.ps1**
  - 安装或更新每日刷新任务：`\AutoTheme\AutoThemeSwitcher-Refresh`。
  - 额外创建补偿任务：`\AutoTheme\AutoThemeSwitcher-Auto`。
  - 触发器：
    - 管理员运行：开机延迟 30 秒。
    - 非管理员运行：当前用户登录时。
    - 每日 00:05 刷新切换计划。
  - 支持 `-RunNow`，安装后立即运行一次刷新任务。

- **ThemeSwitcher.log**
  - 运行日志，自动生成，已加入 `.gitignore`。

- **SunTimesCache.json**
  - 日出、日落缓存，自动生成，已加入 `.gitignore`。
  - 按日期保存缓存，避免今天和次日数据互相覆盖。

## 使用方式

安装或更新计划任务：

```powershell
.\InstallTask.ps1
```

安装后立即刷新一次：

```powershell
.\InstallTask.ps1 -RunNow
```

手动强制切换：

```powershell
.\AutoThemeSwitcher.ps1 -Mode Light
.\AutoThemeSwitcher.ps1 -Mode Dark
```

按当前时间自动判断并补偿：

```powershell
.\AutoThemeSwitcher.ps1 -Mode Auto
```

## 刷新策略

常规 `Light` / `Dark` 切换：

- 写入 Windows 主题注册表。
- 广播 `WM_SETTINGCHANGE`。
- 调用 `UpdatePerUserSystemParameters`。
- 默认不重启 `explorer.exe`。

`Auto` 补偿切换：

- 先判断当前时间应使用浅色还是深色。
- 如果注册表已经是目标值，不重复写入注册表。
- 仍会强制刷新外观。
- 必要时重启 `explorer.exe`，用于同步任务栏主题。

## 运行环境

- Windows 10/11
- PowerShell 5.1+

## 注意事项

- 如需使用系统级「开机触发」，建议以管理员权限运行 `InstallTask.ps1`。
- 非管理员运行时会使用「当前用户登录」触发器。
- 如果无法创建 `\AutoTheme\` 任务文件夹，脚本会回退到任务计划程序根路径 `\`。
- `Auto` 补偿路径可能短暂重启任务栏，这是为了解决任务栏不跟随主题的问题。
- 如需修改坐标或刷新时间，请调整 `AutoThemeSwitcher.ps1`、`UpdateSchedule.ps1` 或 `InstallTask.ps1`。
