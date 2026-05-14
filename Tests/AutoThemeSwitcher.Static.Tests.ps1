$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'AutoThemeSwitcher.ps1'
$source = Get-Content -Raw -Path $scriptPath

function Assert-Contains {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Message
    )

    if ($Text -notmatch $Pattern) {
        throw $Message
    }
}

Assert-Contains `
    -Text $source `
    -Pattern 'function\s+Set-WindowsTheme[\s\S]*?\[switch\]\$ForceRefresh' `
    -Message 'Set-WindowsTheme should expose a ForceRefresh switch.'

Assert-Contains `
    -Text $source `
    -Pattern 'function\s+Invoke-ThemeRefresh[\s\S]*?\[switch\]\$AllowExplorerRestart' `
    -Message 'Invoke-ThemeRefresh should expose an AllowExplorerRestart switch for controlled explorer fallback.'

Assert-Contains `
    -Text $source `
    -Pattern 'Set-WindowsTheme\s+-IsLightMode\s+\$isLightMode\s+-ForceRefresh\s+-AllowExplorerRestart' `
    -Message 'Auto mode should force refresh and allow explorer fallback for startup/logon catch-up.'

Assert-Contains `
    -Text $source `
    -Pattern 'Save-SunriseSunsetCache' `
    -Message 'Successful sunrise/sunset API responses should be cached.'

Assert-Contains `
    -Text $source `
    -Pattern 'Get-CachedSunriseSunset' `
    -Message 'API failure should try cached sunrise/sunset values before fixed fallback times.'

Assert-Contains `
    -Text $source `
    -Pattern '\$dateKey\s*=\s*\$Date\.ToString\("yyyy-MM-dd"\)[\s\S]*?\$cache\[\$dateKey\]' `
    -Message 'Sunrise/sunset cache should be keyed by date so tomorrow lookups do not overwrite today.'

Write-Host 'AutoThemeSwitcher static checks passed.'
