# Windows 自动日出/日落主题切换

基于北京日出/日落时间，在日出时切换浅色、日落时切换深色。通过计划任务自动运行，尽量避免打扰用户（默认不重启 explorer）。

## 文件说明

- **AutoThemeSwitcher.ps1**
  - 实际执行主题切换逻辑。
  - 支持参数：`-Mode Auto|Light|Dark`。
  - 写入注册表：
    - `HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize`
    - `SystemUsesLightTheme` / `AppsUseLightTheme`
  - 日志输出到 `ThemeSwitcher.log`（UTF-8）。

- **UpdateSchedule.ps1**
  - 固定使用北京坐标（39.9042, 116.4074）调用 sunrise-sunset API。
  - 按 **China Standard Time** 计算当日日出/日落时间。
  - 创建/更新两个一次性任务：
    - `\AutoTheme\AutoThemeSwitcher-Light`
    - `\AutoTheme\AutoThemeSwitcher-Dark`
  - 如果当天时间已过，会自动排到次日。

- **InstallTask.ps1**
  - 安装/更新每日刷新任务：`\AutoTheme\AutoThemeSwitcher-Refresh`（若无法创建文件夹则回退到根路径 `\`）。
  - 触发器：
    - 管理员运行：开机延迟 30 秒
    - 非管理员运行：当前用户登录时
    - 每日 00:05 刷新
  - 可选参数 `-RunNow`：立即运行一次刷新。

- **ThemeSwitcher.log**
  - 运行日志（自动生成）。

## 使用方式

1) 安装/更新刷新任务：

```powershell
.
\InstallTask.ps1
```

或立即刷新一次：

```powershell
.
\InstallTask.ps1 -RunNow
```

2) 手动强制切换：

```powershell
.
\AutoThemeSwitcher.ps1 -Mode Light
.
\AutoThemeSwitcher.ps1 -Mode Dark
```

## 运行环境

- Windows 10/11
- PowerShell 5.1+

## 注意事项

- 若希望使用“开机触发”，建议以管理员权限运行 `InstallTask.ps1`。
- 若无法创建 `\AutoTheme\` 任务文件夹，会自动回退到根路径 `\`。

## 备注

- 默认不重启 `explorer.exe`，仅做“尽力刷新”。
- 如需改坐标或刷新时间，请修改 `UpdateSchedule.ps1` / `InstallTask.ps1`。
