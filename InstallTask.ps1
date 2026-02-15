# ============================================================================
# 创建任务计划（每日刷新 + 可选立即运行）
# ============================================================================

param(
    [switch]$RunNow
)

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  Creating AutoThemeSwitcher Task" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

$taskPath = '\AutoTheme\'
$refreshTaskName = "AutoThemeSwitcher-Refresh"
$scriptPath = "$PSScriptRoot\UpdateSchedule.ps1"
$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Check script exists
if (-not (Test-Path $scriptPath)) {
    Write-Host "ERROR: Script not found: $scriptPath" -ForegroundColor Red
    exit
}

# Ensure task folder exists (Task Scheduler COM)
function Ensure-TaskFolder {
    param([string]$Path)

    try {
        $service = New-Object -ComObject "Schedule.Service"
        $service.Connect()
        $root = $service.GetFolder('\')
        $folderName = $Path.Trim('\\')
        if ($folderName) {
            try { $root.GetFolder('\\' + $folderName) | Out-Null }
            catch { $root.CreateFolder($folderName, $null) | Out-Null }
        }

        return $true
    }
    catch {
        Write-Host "WARNING: Failed to create task folder $Path" -ForegroundColor Yellow
        return $false
    }
}

$folderReady = Ensure-TaskFolder -Path $taskPath
if (-not $folderReady) {
    Write-Host "Fallback: use root task path '\'" -ForegroundColor Yellow
    $taskPath = '\'
}

# Remove existing task
$existingTask = Get-ScheduledTask -TaskName $refreshTaskName -TaskPath $taskPath -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $refreshTaskName -TaskPath $taskPath -Confirm:$false
}

# 创建任务操作
$action = New-ScheduledTaskAction `
    -Execute "PowerShell.exe" `
    -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

# Trigger 1: Admin -> startup; Non-admin -> user logon
if ($isElevated) {
    Write-Host "Setting up Trigger 1: At startup (delay 30s)..." -ForegroundColor Yellow
    $trigger1 = New-ScheduledTaskTrigger -AtStartup
    $trigger1.Delay = "PT30S"
}
else {
    Write-Host "Setting up Trigger 1: At logon (current user)..." -ForegroundColor Yellow
    $trigger1 = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERNAME"
}

# 触发器2: 每天 00:05 刷新日出/日落
Write-Host "Setting up Trigger 2: Daily 00:05 refresh..." -ForegroundColor Yellow
$trigger2 = New-ScheduledTaskTrigger -Daily -At ([datetime]::Today.AddMinutes(5))

$triggers = @($trigger1, $trigger2)

# 任务设置
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable:$false `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

# Principal: Admin -> Highest; Non-admin -> Limited
if ($isElevated) {
    $principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive -RunLevel Highest
}
else {
    $principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive -RunLevel Limited
}

# 注册任务
Write-Host "Registering task..." -ForegroundColor Yellow
try {
    Register-ScheduledTask `
        -TaskName $refreshTaskName `
        -TaskPath $taskPath `
        -Action $action `
        -Trigger $triggers `
        -Settings $settings `
        -Principal $principal `
        -Description "Refresh sunrise/sunset schedule daily" `
        -ErrorAction Stop | Out-Null
    
    Write-Host ""
    Write-Host "SUCCESS! Task created successfully!" -ForegroundColor Green
    Write-Host ""
    
    # 验证任务
    $task = Get-ScheduledTask -TaskName $refreshTaskName -TaskPath $taskPath -ErrorAction Stop
    $taskInfo = Get-ScheduledTaskInfo -TaskName $refreshTaskName -TaskPath $taskPath -ErrorAction Stop
    
    Write-Host "Task Details:" -ForegroundColor Cyan
    Write-Host "  - Name: $refreshTaskName"
    Write-Host "  - State: $($task.State)"
    Write-Host "  - Next Run: $($taskInfo.NextRunTime)"
    Write-Host "  - Triggers: $($task.Triggers.Count)"
    Write-Host ""
    
    # List triggers
    Write-Host "Triggers:" -ForegroundColor Cyan
    $triggerIndex = 1
    foreach ($t in $task.Triggers) {
        if ($t.CimClass.CimClassName -eq "MSFT_TaskBootTrigger") {
            Write-Host "  $triggerIndex. At system startup (delay: $($t.Delay))"
        }
        elseif ($t.CimClass.CimClassName -eq "MSFT_TaskLogonTrigger") {
            Write-Host "  $triggerIndex. At user logon"
        }
        elseif ($t.CimClass.CimClassName -eq "MSFT_TaskDailyTrigger") {
            $dailyAt = $t.StartBoundary
            if ($dailyAt -is [datetime]) {
                $dailyAt = $dailyAt.ToString('HH:mm')
            }
            else {
                try { $dailyAt = ([datetime]$dailyAt).ToString('HH:mm') } catch {}
            }
            Write-Host "  $triggerIndex. Daily at $dailyAt"
        }
        $triggerIndex++
    }
    
    Write-Host ""
    Write-Host "Verification:" -ForegroundColor Cyan
    Write-Host "  1. Open Task Scheduler (Win+R -> taskschd.msc)"
    Write-Host "  2. Find task '$taskPath$refreshTaskName'"
    Write-Host "  3. Check 'Triggers' tab"
    Write-Host "  4. Check log: e:\auto_theme_change\ThemeSwitcher.log"
    Write-Host ""
    
    if ($RunNow) {
        Write-Host ""
        Write-Host "Running task..." -ForegroundColor Yellow
        Start-ScheduledTask -TaskName $refreshTaskName -TaskPath $taskPath
        Start-Sleep -Seconds 3
        Write-Host ""
        Write-Host "Check the log file for results:" -ForegroundColor Green
        Write-Host "  $PSScriptRoot\ThemeSwitcher.log"
    }
    
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    if (-not $isElevated) {
        Write-Host "Tip: Admin is only required for startup-level (AtStartup) task." -ForegroundColor Yellow
    }
}

Write-Host ""
