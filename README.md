# KMTECH MRP — 생산·재고·BOM 통합 관리 시스템

소규모 제조업체(5인 이하, 단일 PC)를 위한 경량 MRP(Material Requirements Planning) 시스템입니다.
별도 설치 없이 실행 파일 하나로 동작하며, 모든 데이터는 Excel 파일에 저장됩니다.

> 현재 버전: **v1.3** (2026-09-20)

---

## 주요 기능

| 기능 | 설명 |
|------|------|
| 📦 **품목 기준정보** | 자재/부품 코드, 기준재고·안전재고·목표재고·리드타임 관리 |
| ⚙️ **BOM 관리** | 완제품 → 부품 투입 구조 정규형 편집 |
| 🏭 **생산 실적관리** | 일자별 생산 로그(투입/양품/불량) 입력·수정·삭제 |
| 🔄 **자재 입출고** | 입고/출고 원장 관리, 실시간 재고 계산 |
| 📦 **발주 관리** | 미입고 발주 추적, 일괄 출력(전표/라벨) |
| 🚚 **완제품 출하관리** | 출하/판매 기록, 완제품 재고 자동 산출 |
| 📊 **KPI 대시보드** | 재고/생산능력/수요시뮬레이션 8종 지표 |
| 💾 **자동 백업** | 저장 시 + 매일 00:00 자동 백업, 7일 보관 |

## 화면 구성

- **KPI 대시보드** — 재고 현황 · 생산능력 · 수요 시뮬레이션을 한 화면에 표시
- **완제품 탭** — 완제품 기준정보 · BOM · 생산실적 · 출하관리
- **자재 탭** — 자재 기준정보 · 입출고 원장 · 발주관리
- 각 데이터는 모달에서 즉시 입력/수정/삭제 가능

## 설치 및 실행

### 방법 1 — 배포 실행 파일 (권장, 별도 설치 불필요)

1. [Releases](https://github.com/mr-leeeee/kmtech1-mrp/releases)에서 최신 `KMTECH_MRP_v1.3.zip` 다운로드
2. 압축 해제 후 `KMTECH_MRP.exe` 더블클릭
3. 브라우저에서 `http://127.0.0.1:8000` 접속
4. 최초 실행 시 exe와 같은 폴더에 `shelf_MRP_개선판.xlsx`가 자동 생성됩니다 (모든 데이터는 이 파일에 저장)

### 방법 2 — 소스 실행 (개발자)

```bash
# Python 3.10+ 필요
pip install fastapi uvicorn openpyxl pydantic

python server.py
# → http://127.0.0.1:8000
```

### 테스트

```bash
python -m pytest test_validation.py test_crud_system.py
# 24 passed
```

## 데이터 & 백업

- 마스터 데이터: `shelf_MRP_개선판.xlsx` (실행 폴더)
- 백업 위치: `backup/YYYYMMDD/*.bak` — 저장할 때마다 + 매일 00:00 자동 생성, 7일 보관
- 복구 절차: [백업/복구 가이드](docs/backup_restore.md)

## 기술 스택

- **백엔드**: Python 3.10+ · FastAPI · Uvicorn
- **데이터**: Excel (openpyxl)
- **프론트엔드**: Vanilla HTML/CSS/JS (빌드 불필요)
- **배포**: PyInstaller 단일 실행파일

## 프로젝트 구조

```
├── server.py            # FastAPI 서버 (API + 정적 파일)
├── mrp_engine.py        # MRP 핵심 엔진 (재고/생산/수요 계산)
├── excel_sync.py        # Excel 로드/저장/백업 동기화
├── static/index.html    # 웹 UI (단일 파일)
├── docs/                # 사용설명서·백업복구 가이드
└── test_*.py            # pytest 테스트
```

## 라이선스

사내/소규모 사업장 내부 사용을 목적으로 제작되었습니다. 무단 영리 배포 금지.

---

문의: mr-leeeee (GitHub Issues 활용)