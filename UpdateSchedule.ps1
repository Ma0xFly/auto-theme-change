# ============================================================================
# UpdateSchedule.ps1
# Compute Beijing sunrise/sunset and update Light/Dark tasks
# ============================================================================

$LogFile = Join-Path $PSScriptRoot 'ThemeSwitcher.log'
$MaxLogSize = 1MB
$BeijingLatitude = 39.9042
$BeijingLongitude = 116.4074
$ChinaTimeZoneId = 'China Standard Time'
$Backslash = [char]92

$TaskPath = ([string]$Backslash) + 'AutoTheme' + ([string]$Backslash)
$LightTaskName = 'AutoThemeSwitcher-Light'
$DarkTaskName = 'AutoThemeSwitcher-Dark'

$ThemeScriptPath = Join-Path $PSScriptRoot 'AutoThemeSwitcher.ps1'

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] $Message"

    if (Test-Path $LogFile) {
        if ((Get-Item $LogFile).Length -gt $MaxLogSize) {
            Remove-Item $LogFile -Force
        }
    }

    Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8
    Write-Host $logMessage
}

function Get-ChinaTimeZone {
    try {
        return [TimeZoneInfo]::FindSystemTimeZoneById($ChinaTimeZoneId)
    }
    catch {
        Write-Log "Warning: Could not resolve China Standard Time, using local time zone"
        return [TimeZoneInfo]::Local
    }
}

function Ensure-TaskFolder {
    param([string]$Path)

    try {
        $service = New-Object -ComObject 'Schedule.Service'
        $service.Connect()
        $root = $service.GetFolder([string]$Backslash)
        $folderName = $Path.Trim($Backslash)
        if ($folderName) {
            try { $root.GetFolder(([string]$Backslash) + $folderName) | Out-Null }
            catch { $root.CreateFolder($folderName, $null) | Out-Null }
        }
    }
    catch {
        Write-Log "Warning: Could not create task folder $Path"
    }
}

function Get-SunriseSunsetForDate {
    param(
        [double]$Latitude,
        [double]$Longitude,
        [DateTime]$Date,
        [TimeZoneInfo]$TimeZone
    )

    try {
        $baseUri = 'https://api.sunrise-sunset.org/json'
        $params = @{
            lat       = $Latitude
            lng       = $Longitude
            formatted = 0
            date      = $Date.ToString('yyyy-MM-dd')
        }
        $response = Invoke-RestMethod -Uri $baseUri -Body $params -Method Get -TimeoutSec 10

        if ($response.status -eq 'OK') {
            $sunriseUtc = [DateTimeOffset]::Parse($response.results.sunrise).UtcDateTime
            $sunsetUtc = [DateTimeOffset]::Parse($response.results.sunset).UtcDateTime

            $sunriseLocal = [TimeZoneInfo]::ConvertTimeFromUtc($sunriseUtc, $TimeZone)
            $sunsetLocal = [TimeZoneInfo]::ConvertTimeFromUtc($sunsetUtc, $TimeZone)

            return @{
                Sunrise = $sunriseLocal
                Sunset  = $sunsetLocal
            }
        }

        Write-Log "Warning: Sunrise-sunset API returned non-OK status"
        return $null
    }
    catch {
        Write-Log "Error: Failed to fetch sunrise/sunset time - $($_.Exception.Message)"
        return $null
    }
}

function Get-FallbackTimes {
    param([DateTime]$CurrentTime)

    $baseDate = $CurrentTime.Date
    $fallbackSunrise = $baseDate.AddHours(7)
    $fallbackSunset = $baseDate.AddHours(18)

    if ($fallbackSunrise -le $CurrentTime) {
        $fallbackSunrise = $fallbackSunrise.AddDays(1)
    }
    if ($fallbackSunset -le $CurrentTime) {
        $fallbackSunset = $fallbackSunset.AddDays(1)
    }

    return @{
        Sunrise = $fallbackSunrise
        Sunset  = $fallbackSunset
    }
}

function Register-OneShotTask {
    param(
        [string]$TaskName,
        [DateTime]$RunAt,
        [string]$Mode
    )

    $argument = '-ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $ThemeScriptPath + '" -Mode ' + $Mode

    $action = New-ScheduledTaskAction `
        -Execute 'PowerShell.exe' `
        -Argument $argument

    $trigger = New-ScheduledTaskTrigger -Once -At $RunAt

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -RunOnlyIfNetworkAvailable:$false `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1)

    $principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive -RunLevel Limited

    try {
        Register-ScheduledTask `
            -TaskName $TaskName `
            -TaskPath $TaskPath `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Principal $principal `
            -Description "Auto switch Windows theme at next $Mode time" `
            -Force | Out-Null

        return $true
    }
    catch {
        Write-Log ("Error: Failed to register task '" + $TaskName + "' - " + $_.Exception.Message)
        return $false
    }
}

Write-Log '========================================='
Write-Log 'Start updating sunrise/sunset trigger tasks'
Write-Log '========================================='

if (-not (Test-Path $ThemeScriptPath)) {
    Write-Log "Error: Theme switcher script not found: $ThemeScriptPath"
    exit 1
}

$timeZone = Get-ChinaTimeZone
$utcNow = (Get-Date).ToUniversalTime()
$now = [TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $timeZone)

Ensure-TaskFolder -Path $TaskPath

$todayTimes = Get-SunriseSunsetForDate -Latitude $BeijingLatitude -Longitude $BeijingLongitude -Date $now -TimeZone $timeZone

if ($null -eq $todayTimes) {
    Write-Log 'Using fallback times: 07:00 / 18:00'
    $nextTimes = Get-FallbackTimes -CurrentTime $now
    $lightAt = $nextTimes.Sunrise
    $darkAt = $nextTimes.Sunset
}
else {
    $lightAt = $todayTimes.Sunrise
    $darkAt = $todayTimes.Sunset

    $tomorrowTimes = $null

    if ($lightAt -le $now) {
        $tomorrowTimes = Get-SunriseSunsetForDate -Latitude $BeijingLatitude -Longitude $BeijingLongitude -Date $now.AddDays(1) -TimeZone $timeZone
        if ($null -ne $tomorrowTimes) {
            $lightAt = $tomorrowTimes.Sunrise
        }
    }

    if ($darkAt -le $now) {
        if ($null -eq $tomorrowTimes) {
            $tomorrowTimes = Get-SunriseSunsetForDate -Latitude $BeijingLatitude -Longitude $BeijingLongitude -Date $now.AddDays(1) -TimeZone $timeZone
        }
        if ($null -ne $tomorrowTimes) {
            $darkAt = $tomorrowTimes.Sunset
        }
    }
}

Write-Log ('Planned Light task time: ' + $lightAt.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Log ('Planned Dark task time: ' + $darkAt.ToString('yyyy-MM-dd HH:mm:ss'))

$lightTaskResult = Register-OneShotTask -TaskName $LightTaskName -RunAt $lightAt -Mode 'Light'
$darkTaskResult = Register-OneShotTask -TaskName $DarkTaskName -RunAt $darkAt -Mode 'Dark'

Write-Log '========================================='
if ($lightTaskResult -and $darkTaskResult) {
    Write-Log 'Sunrise/sunset tasks updated'
    Write-Log '========================================='
    exit 0
}

Write-Log 'Task update completed with errors'
Write-Log 'Tip: if access is denied, run PowerShell as Administrator once and retry'
Write-Log '========================================='
exit 1
