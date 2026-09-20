# KMTECH MRP 실무 전환 설계서

## 1. 현재 상태
- 단일 Excel 파일 기반 MRP
- FastAPI + Uvicorn 단일 인스턴스
- 개인 PC localhost:8000
- 인증 없음, 로깅 파일 무한 증가

## 2. 마이그레이션 목표
- 데이터 무결성 및 동시성 보장
- 다중 사용자 지원
- 운영 모니터링 및 백업 정책

## 3. 데이터베이스 전환
### 3.1 SQLite → PostgreSQL
**단계 1: 스키마 설계**
- tables: items, bom, inventory, production_logs, fg_transactions, purchase_orders, users, audit_log
- 관계형 키, 트랜잭션 보장

**단계 2: 마이그레이션 스크립트**
- excel_sync.py → db_sync.py
- openpyxl 로드 → SQLAlchemy ORM
- 초기 로드: Excel → PostgreSQL 일괄 import

**단계 3: 애플리케이션 수정**
- mrp_engine.py의 in-memory 리스트 → DB 쿼리
- 캐시 전략: Redis optional

### 3.2 인증/인가
- FastAPI Users + JWT
- 역할: admin, planner, viewer
- 미들웨어로 엔드포인트 보호

## 4. 로깅 로테이션
- logging.handlers.TimedRotatingFileHandler
- 일별 회전, 30일 보관
- 로그 레벨 분리: access, error

## 5. 운영 자동화
- start_mrp.ps1 → systemd / Docker Compose
- 백업: PostgreSQL pg_dump 매일 02:00
- 모니터링: health check /api/health

## 6. 일정
- Week1: 스키마 설계 및 마이그레이션 스크립트
- Week2: 인증 구현 및 기존 기능 재현
- Week3: 테스트 및 운영 배포

## 7. 위험 요소
- Excel 수식 의존성 제거 필요
- 기존 사용자 교육
