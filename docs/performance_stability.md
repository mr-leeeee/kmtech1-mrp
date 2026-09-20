# 성능 및 안정성 개선안

## 성능
1. 캐시 도입
   - 재고 계산 결과 5분 캐시
   - appState에 타임스탬프 기반 무효화

2. 비동기 IO
   - Excel 쓰기 대량 작업 시 asyncio.to_thread 사용

3. DB 전환
   - Excel → PostgreSQL로 전환 시 쿼리 인덱싱

## 안정성
1. 에러 처리
   - 모든 엔드포인트 try/except + 상세 로그
   - 글로벌 예외 핸들러

2. 백업
   - Excel 백업 7일 로테이션 적용 완료
   - start_mrp.ps1로 중복 실행 방지

3. 모니터링
   - /api/health 엔드포인트 추가
   - 로그 로테이션 TimedRotatingFileHandler

4. 데이터 무결성
   - 쓰기 전 원자적 백업
   - ExcelLock 컨텍스트 매니저 유지
