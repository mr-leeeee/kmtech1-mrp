# KMTECH MRP 자동 시작 스크립트
# 중복 프로세스 정리 -> 서버 기동

$projectPath = "C:\vibecode\kmtech1"
$port = 8000

Write-Host "=== KMTECH MRP 시작 점검 ===" -ForegroundColor Cyan

# 포트 점유 확인
$listener = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($listener) {
    Write-Host "포트 $port 점유 중 감지. 기존 Python 프로세스 종료 시도..." -ForegroundColor Yellow
    Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
}

# 잔여 프로세스 정리
Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Write-Host "프로젝트 폴더: $projectPath" -ForegroundColor Green
Set-Location $projectPath

# 서버 시작
Write-Host "서버 기동 중..." -ForegroundColor Green
Start-Process python -ArgumentList "server.py" -WorkingDirectory $projectPath -WindowStyle Hidden

Start-Sleep -Seconds 3

# 기동 확인
$listener = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($listener) {
    Write-Host "✅ 서버 정상 기동: http://127.0.0.1:$port" -ForegroundColor Green
} else {
    Write-Host "⚠️ 서버 기동 확인 실패. 로그 확인 필요." -ForegroundColor Red
}
