# KMTECH MRP 빌드 패키지 생성
$src = "C:\vibecode\kmtech1"
$dist = "C:\vibecode\kmtech1\dist\KMTECH_MRP_v1.1"

Remove-Item $dist -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $dist | Out-Null

Copy-Item "$src\server.py" $dist
Copy-Item "$src\mrp_engine.py" $dist
Copy-Item "$src\excel_sync.py" $dist
Copy-Item "$src\static" $dist -Recurse
Copy-Item "$src\start_mrp.ps1" $dist
Copy-Item "$src\docs" $dist -Recurse

Write-Host "패키지 생성 완료: $dist" -ForegroundColor Green
Write-Host "실행 방법: start_mrp.ps1 실행 후 http://127.0.0.1:8000"
