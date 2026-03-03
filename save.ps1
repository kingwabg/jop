param(
    [string]$msg = "Update: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
)

Write-Host "🚀 깃허브 저장 시작..." -ForegroundColor Cyan

# 1. 변경된 모든 파일 추가
git add .
if ($LASTEXITCODE -ne 0) { Write-Host "❌ git add 실패" -ForegroundColor Red; exit }

# 2. 커밋 메시지 작성 및 커밋
git commit -m "$msg"
if ($LASTEXITCODE -ne 0) { Write-Host "⚠️ 커밋할 내용이 없거나 이미 최신 상태입니다." -ForegroundColor Yellow }

# 3. 깃허브 서버로 전송
git push
if ($LASTEXITCODE -ne 0) { Write-Host "❌ git push 실패" -ForegroundColor Red; exit }

Write-Host "✅ 저장 완료! (메시지: $msg)" -ForegroundColor Green
