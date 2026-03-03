param([string]$Action, [string]$Msg)

if ($Action -ne "저장") {
    Write-Host "Usage: .\깃 저장 'message'"
    exit
}

$commitMsg = if ([string]::IsNullOrWhiteSpace($Msg)) { "Update: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" } else { $Msg }

Write-Host ">>> Git Add..."
git add .

Write-Host ">>> Git Commit..."
git commit -m $commitMsg

Write-Host ">>> Git Push..."
git push

Write-Host ">>> Done!"
