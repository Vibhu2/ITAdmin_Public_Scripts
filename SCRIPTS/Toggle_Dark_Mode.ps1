# Windows 11 Dark Mode / Light Mode Toggle Script
# By default toggles ALL three settings together.
# Use switches to selectively control individual settings.
#
# Usage:
#   .\toggle-dark-mode.ps1                          # Toggle all (default)
#   .\toggle-dark-mode.ps1 -SkipApps               # Skip AppsUseLightTheme
#   .\toggle-dark-mode.ps1 -SkipSystem              # Skip SystemUsesLightTheme
#   .\toggle-dark-mode.ps1 -SkipAppsDarkMode        # Skip AppsUseDarkMode
#   .\toggle-dark-mode.ps1 -SkipSystem -SkipAppsDarkMode  # Only toggle Apps
#   .\toggle-dark-mode.ps1 -ForceDark               # Force dark mode regardless
#   .\toggle-dark-mode.ps1 -ForceLight              # Force light mode regardless

param(
    [switch]$SkipApps,          # Skip toggling AppsUseLightTheme
    [switch]$SkipSystem,        # Skip toggling SystemUsesLightTheme
    [switch]$SkipAppsDarkMode,  # Skip toggling AppsUseDarkMode
    [switch]$ForceDark,         # Force dark mode (overrides toggle)
    [switch]$ForceLight         # Force light mode (overrides toggle)
)

$RegistryPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"

try {
    # Determine target value
    $CurrentAppsLight = (Get-ItemProperty -Path $RegistryPath -Name AppsUseLightTheme -ErrorAction SilentlyContinue).AppsUseLightTheme

    if ($ForceDark) {
        $NewValue = 0   # Dark
    } elseif ($ForceLight) {
        $NewValue = 1   # Light
    } else {
        # Toggle based on current state
        $NewValue = if ($CurrentAppsLight -eq 0) { 1 } else { 0 }
    }

    # Inverse for AppsUseDarkMode (1=Dark, 0=Light)
    $DarkModeValue = if ($NewValue -eq 0) { 1 } else { 0 }

    $Mode = if ($NewValue -eq 0) { "DARK MODE" } else { "LIGHT MODE" }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Switching to $Mode" -ForegroundColor $(if ($NewValue -eq 0) { "Blue" } else { "Yellow" })
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Registry Updates:" -ForegroundColor Yellow

    # AppsUseLightTheme
    if (-not $SkipApps) {
        Set-ItemProperty -Path $RegistryPath -Name AppsUseLightTheme -Value $NewValue -Type DWord -Force
        Write-Host "  [ON]  AppsUseLightTheme    → $NewValue" -ForegroundColor Green
    } else {
        Write-Host "  [--]  AppsUseLightTheme    → skipped" -ForegroundColor DarkGray
    }

    # SystemUsesLightTheme
    if (-not $SkipSystem) {
        Set-ItemProperty -Path $RegistryPath -Name SystemUsesLightTheme -Value $NewValue -Type DWord -Force
        Write-Host "  [ON]  SystemUsesLightTheme → $NewValue" -ForegroundColor Green
    } else {
        Write-Host "  [--]  SystemUsesLightTheme → skipped" -ForegroundColor DarkGray
    }

    # AppsUseDarkMode
    if (-not $SkipAppsDarkMode) {
        Set-ItemProperty -Path $RegistryPath -Name AppsUseDarkMode -Value $DarkModeValue -Type DWord -Force -ErrorAction SilentlyContinue
        Write-Host "  [ON]  AppsUseDarkMode      → $DarkModeValue" -ForegroundColor Green
    } else {
        Write-Host "  [--]  AppsUseDarkMode      → skipped" -ForegroundColor DarkGray
    }

    Write-Host ""

    # Restart File Explorer to apply changes
    Write-Host "Restarting File Explorer..." -ForegroundColor Cyan
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
    Start-Process explorer.exe

    Write-Host "✓ Done!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

} catch {
    Write-Host "✗ Error: $_" -ForegroundColor Red
    exit 1
}
