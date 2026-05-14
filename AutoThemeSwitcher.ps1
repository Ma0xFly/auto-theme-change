# ============================================================================
# AutoThemeSwitcher.ps1
# 智能 Windows 主题切换脚本 - 基于地理位置的日出日落时间
# ============================================================================

# 参数
param(
    [ValidateSet("Auto", "Light", "Dark")]
    [string]$Mode = "Auto"
)

# 配置参数
$LogFile = "$PSScriptRoot\ThemeSwitcher.log"
$SunTimesCacheFile = "$PSScriptRoot\SunTimesCache.json"
$MaxLogSize = 1MB
$BeijingLatitude = 39.9042
$BeijingLongitude = 116.4074
$ChinaTimeZoneId = "China Standard Time"

# 日志函数
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # 控制日志文件大小
    if (Test-Path $LogFile) {
        if ((Get-Item $LogFile).Length -gt $MaxLogSize) {
            Remove-Item $LogFile -Force
        }
    }
    
    Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8
    Write-Host $logMessage
}

# 获取中国标准时间时区
function Get-ChinaTimeZone {
    try {
        return [TimeZoneInfo]::FindSystemTimeZoneById($ChinaTimeZoneId)
    }
    catch {
        Write-Log "警告: 无法获取 China Standard Time，使用本地时区"
        return [TimeZoneInfo]::Local
    }
}

function Save-SunriseSunsetCache {
    param(
        [DateTime]$Date,
        [DateTime]$Sunrise,
        [DateTime]$Sunset,
        [string]$TimeZoneId
    )

    try {
        $dateKey = $Date.ToString("yyyy-MM-dd")
        $cache = [ordered]@{}

        if (Test-Path $SunTimesCacheFile) {
            $existingCache = Get-Content -Raw -Path $SunTimesCacheFile | ConvertFrom-Json
            foreach ($property in $existingCache.PSObject.Properties) {
                $cache[$property.Name] = $property.Value
            }
        }

        $cache[$dateKey] = [ordered]@{
            Sunrise    = $Sunrise.ToString("o")
            Sunset     = $Sunset.ToString("o")
            TimeZoneId = $TimeZoneId
            UpdatedAt  = (Get-Date).ToString("o")
        }

        $cache | ConvertTo-Json | Set-Content -Path $SunTimesCacheFile -Encoding UTF8
        Write-Log "已缓存 $dateKey 日出日落时间"
    }
    catch {
        Write-Log "警告: 无法写入日出日落缓存 - $($_.Exception.Message)"
    }
}

function Get-CachedSunriseSunset {
    param(
        [DateTime]$Date,
        [TimeZoneInfo]$TimeZone
    )

    try {
        $dateKey = $Date.ToString("yyyy-MM-dd")

        if (-not (Test-Path $SunTimesCacheFile)) {
            Write-Log "未找到日出日落缓存"
            return $null
        }

        $cache = Get-Content -Raw -Path $SunTimesCacheFile | ConvertFrom-Json
        $entry = $cache.PSObject.Properties[$dateKey].Value
        if ($null -eq $entry) {
            Write-Log "未找到 $dateKey 的日出日落缓存"
            return $null
        }

        if ($entry.TimeZoneId -and $entry.TimeZoneId -ne $TimeZone.Id) {
            Write-Log "日出日落缓存时区不匹配，跳过缓存"
            return $null
        }

        $sunrise = [DateTime]::Parse($entry.Sunrise)
        $sunset = [DateTime]::Parse($entry.Sunset)

        Write-Log "使用缓存日出日落时间: 日出 $($sunrise.ToString('HH:mm:ss'))，日落 $($sunset.ToString('HH:mm:ss'))"
        return @{
            Sunrise = $sunrise
            Sunset  = $sunset
        }
    }
    catch {
        Write-Log "警告: 无法读取日出日落缓存 - $($_.Exception.Message)"
        return $null
    }
}

# 获取日出日落时间
function Get-SunriseSunset {
    param(
        [double]$Latitude,
        [double]$Longitude,
        [DateTime]$Date,
        [TimeZoneInfo]$TimeZone
    )
    
    try {
        Write-Log "正在获取日出日落时间..."
        $baseUri = "https://api.sunrise-sunset.org/json"
        $params = @{
            lat       = $Latitude
            lng       = $Longitude
            formatted = 0
            date      = $Date.ToString("yyyy-MM-dd")
        }
        $response = Invoke-RestMethod -Uri $baseUri -Body $params -Method Get -TimeoutSec 10
        
        if ($response.status -eq "OK") {
            $sunriseUtc = [DateTimeOffset]::Parse($response.results.sunrise).UtcDateTime
            $sunsetUtc = [DateTimeOffset]::Parse($response.results.sunset).UtcDateTime
            
            # 转换为指定时区时间（北京时间）
            $sunriseLocal = [TimeZoneInfo]::ConvertTimeFromUtc($sunriseUtc, $TimeZone)
            $sunsetLocal = [TimeZoneInfo]::ConvertTimeFromUtc($sunsetUtc, $TimeZone)
            
            Write-Log "日出时间: $($sunriseLocal.ToString('HH:mm:ss'))"
            Write-Log "日落时间: $($sunsetLocal.ToString('HH:mm:ss'))"
            Save-SunriseSunsetCache -Date $Date -Sunrise $sunriseLocal -Sunset $sunsetLocal -TimeZoneId $TimeZone.Id
            
            return @{
                Sunrise = $sunriseLocal
                Sunset  = $sunsetLocal
            }
        }
        else {
            Write-Log "警告: 日出日落API返回失败状态"
            return $null
        }
    }
    catch {
        Write-Log "错误: 无法获取日出日落时间 - $($_.Exception.Message)"
        return $null
    }
}

# 尽力刷新主题；强刷新场景允许重启资源管理器作为任务栏兜底
function Invoke-ThemeRefresh {
    param([switch]$AllowExplorerRestart)

    try {
        $signature = @"
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    public static extern IntPtr SendMessageTimeout(
        IntPtr hWnd, int Msg, IntPtr wParam, string lParam,
        int fuFlags, int uTimeout, out IntPtr lpdwResult);
}
"@

        if (-not ("NativeMethods" -as [Type])) {
            Add-Type -TypeDefinition $signature -ErrorAction Stop
        }

        $HWND_BROADCAST = [IntPtr]0xffff
        $WM_SETTINGCHANGE = 0x1A
        $SMTO_ABORTIFHUNG = 0x0002
        $result = [IntPtr]::Zero

        [void][NativeMethods]::SendMessageTimeout(
            $HWND_BROADCAST,
            $WM_SETTINGCHANGE,
            [IntPtr]::Zero,
            "ImmersiveColorSet",
            $SMTO_ABORTIFHUNG,
            1000,
            [ref]$result
        )

        Start-Process -FilePath "$env:SystemRoot\System32\rundll32.exe" -ArgumentList "user32.dll,UpdatePerUserSystemParameters" -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-Log "已尝试刷新主题设置（广播设置变更）"

        if ($AllowExplorerRestart) {
            Write-Log "强刷新: 重启 explorer.exe 以同步任务栏主题"
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Start-Process -FilePath "$env:SystemRoot\explorer.exe" -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Log "警告: 刷新主题设置失败 - $($_.Exception.Message)"
    }
}

# 获取当前主题状态
function Get-CurrentTheme {
    try {
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $systemTheme = Get-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -ErrorAction SilentlyContinue
        $appsTheme = Get-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
        
        if ($null -ne $systemTheme -and $null -ne $appsTheme) {
            return @{
                System = $systemTheme.SystemUsesLightTheme
                Apps   = $appsTheme.AppsUseLightTheme
            }
        }
        return $null
    }
    catch {
        Write-Log "错误: 无法读取当前主题状态 - $($_.Exception.Message)"
        return $null
    }
}

# 设置 Windows 主题
function Set-WindowsTheme {
    param(
        [bool]$IsLightMode,
        [switch]$ForceRefresh,
        [switch]$AllowExplorerRestart
    )
    
    $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    $themeValue = if ($IsLightMode) { 1 } else { 0 }
    $themeName = if ($IsLightMode) { "浅色" } else { "深色" }
    
    # 获取当前状态
    $currentTheme = Get-CurrentTheme
    
    $shouldWriteTheme = $true

    # 防抖动: 检查是否需要写入
    if ($null -ne $currentTheme) {
        if ($currentTheme.System -eq $themeValue -and $currentTheme.Apps -eq $themeValue) {
            $shouldWriteTheme = $false
            Write-Log "当前已经是 $themeName 模式，无需写入注册表"
        }
    }
    
    try {
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }

        if ($shouldWriteTheme) {
            Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value $themeValue -Type DWord
            Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value $themeValue -Type DWord
            Write-Log "✓ 成功切换到 $themeName 模式"
            Invoke-ThemeRefresh -AllowExplorerRestart:$AllowExplorerRestart
            return
        }

        if ($ForceRefresh) {
            Write-Log "强制刷新当前 $themeName 模式外观"
            Invoke-ThemeRefresh -AllowExplorerRestart:$AllowExplorerRestart
        }
    }
    catch {
        Write-Log "错误: 无法设置主题 - $($_.Exception.Message)"
    }
}

# 判断应该使用什么主题
function Get-RecommendedTheme {
    param(
        [DateTime]$Sunrise,
        [DateTime]$Sunset,
        [DateTime]$CurrentTime
    )
    
    $currentTimeOnly = $CurrentTime.TimeOfDay
    $sunriseTimeOnly = $Sunrise.TimeOfDay
    $sunsetTimeOnly = $Sunset.TimeOfDay
    
    Write-Log "当前时间: $($CurrentTime.ToString('HH:mm:ss'))"
    
    # PowerShell requires else on same line as closing brace
    if ($currentTimeOnly -ge $sunriseTimeOnly -and $currentTimeOnly -lt $sunsetTimeOnly) {
        Write-Log "判断: 在日出($($Sunrise.ToString('HH:mm')))和日落($($Sunset.ToString('HH:mm')))之间 -> 浅色模式"
        return $true
    }
    else {
        Write-Log "判断: 在日落之后或日出之前 -> 深色模式"
        return $false
    }
}

# 使用保底逻辑（当 API 失败时）
function Use-FallbackLogic {
    param(
        [DateTime]$CurrentTime,
        [switch]$ForceRefresh,
        [switch]$AllowExplorerRestart
    )
    
    Write-Log "使用保底逻辑: 7:00 - 18:00 为浅色模式"
    
    $baseDate = $CurrentTime.Date
    $fallbackSunrise = $baseDate.AddHours(7)
    $fallbackSunset = $baseDate.AddHours(18)
    
    $isLightMode = Get-RecommendedTheme -Sunrise $fallbackSunrise -Sunset $fallbackSunset -CurrentTime $CurrentTime
    Set-WindowsTheme -IsLightMode $isLightMode -ForceRefresh:$ForceRefresh -AllowExplorerRestart:$AllowExplorerRestart
}

# ============================================================================
# 主执行逻辑
# ============================================================================

Write-Log "========================================="
Write-Log "开始执行主题切换脚本 (Mode: $Mode)"
Write-Log "========================================="

$timeZone = Get-ChinaTimeZone
$utcNow = (Get-Date).ToUniversalTime()
$currentTime = [TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $timeZone)

if ($Mode -eq "Light") {
    Set-WindowsTheme -IsLightMode $true
    exit
}

if ($Mode -eq "Dark") {
    Set-WindowsTheme -IsLightMode $false
    exit
}

# Auto 模式：使用北京经纬度 + 中国标准时间
$sunTimes = Get-SunriseSunset -Latitude $BeijingLatitude -Longitude $BeijingLongitude -Date $currentTime -TimeZone $timeZone

if ($null -eq $sunTimes) {
    Write-Log "无法获取日出日落时间，尝试使用缓存"
    $sunTimes = Get-CachedSunriseSunset -Date $currentTime -TimeZone $timeZone

    if ($null -eq $sunTimes) {
        Write-Log "无法使用缓存日出日落时间，使用保底逻辑"
        Use-FallbackLogic -CurrentTime $currentTime -ForceRefresh -AllowExplorerRestart
        exit
    }
}

$isLightMode = Get-RecommendedTheme -Sunrise $sunTimes.Sunrise -Sunset $sunTimes.Sunset -CurrentTime $currentTime
Set-WindowsTheme -IsLightMode $isLightMode -ForceRefresh -AllowExplorerRestart

Write-Log "========================================="
Write-Log "脚本执行完成"
Write-Log "========================================="
