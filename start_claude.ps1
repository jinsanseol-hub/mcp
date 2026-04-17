$app = Get-StartApps | Where-Object { $_.Name -like '*Claude*' }
$app | Format-Table Name, AppID
if ($app) {
    Start-Process ("shell:AppsFolder\" + $app[0].AppID)
    Write-Host "Claude launched: $($app[0].AppID)"
} else {
    Write-Host "Claude not found in Start Apps"
}
