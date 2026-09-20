import os
import logging
import atexit
import time
import traceback
import webbrowser
import threading
from datetime import datetime, timedelta
from typing import Dict, Any, Optional
from fastapi import FastAPI, HTTPException, Request
from fastapi.exceptions import RequestValidationError
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, field_validator
import uvicorn

import excel_sync
from mrp_engine import (MRPSystem, Item, FinishedGood, BOMItem, ProductionLog,
                        Transaction, PurchaseOrder, FGTransaction)

log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mrp_server.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(log_file, encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

# Excel 파일 경로: 환경변수 MRP_EXCEL_FILE 우선, 없으면 기본값 사용
import sys, shutil
def get_base_path():
    if getattr(sys, 'frozen', False):
        # PyInstaller로 빌드된 경우
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return base_path

base_path = get_base_path()

def get_excel_path():
    env_path = os.environ.get("MRP_EXCEL_FILE")
    if env_path:
        return env_path
    if getattr(sys, 'frozen', False):
        # 실행 파일과 같은 폴더에 저장/불러오기
        exe_dir = os.path.dirname(sys.executable)
        excel_path = os.path.join(exe_dir, "shelf_MRP_개선판.xlsx")
        # 처음 실행 시 번들에서 복사
        bundled_path = os.path.join(base_path, "shelf_MRP_개선판.xlsx")
        if not os.path.exists(excel_path) and os.path.exists(bundled_path):
            try:
                shutil.copy2(bundled_path, excel_path)
                logger.info("Excel 번들 파일을 실행 폴더로 복사: %s", excel_path)
            except Exception as e:
                logger.warning("Excel 파일 복사 실패: %s", e)
        return excel_path
    else:
        return os.path.join(base_path, "shelf_MRP_개선판.xlsx")

EXCEL_FILE = get_excel_path()
logger.info("Excel 파일 경로: %s", EXCEL_FILE)

app = FastAPI(title="MRP 생산·재고·BOM 통합 관리 시스템", version="1.1.0")

# 개발/운영 모드 분리: MRP_DEBUG=1 일 때만 상세 에러 노출
DEBUG_MODE = os.environ.get("MRP_DEBUG", "0") == "1"

@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    return JSONResponse(
        status_code=exc.status_code,
        content={"success": False, "message": str(exc.detail)},
    )

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    logger.warning("요청 검증 실패: %s %s → %s", request.method, request.url.path, exc.errors())
    errors = []
    for err in exc.errors():
        loc = ".".join(str(p) for p in err.get("loc", []) if p != "body")
        errors.append({"field": loc, "message": err.get("msg", "")})
    return JSONResponse(
        status_code=400,
        content={"success": False, "message": "입력값이 올바르지 않습니다", "errors": errors},
    )

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.error("Unhandled error: %s %s → %s\n%s", request.method, request.url.path, exc, traceback.format_exc())
    msg = f"서버 내부 오류: {exc}" if DEBUG_MODE else "서버 내부 오류가 발생했습니다"
    return JSONResponse(status_code=500, content={"success": False, "message": msg})

# Global MRP instance
mrp: MRPSystem = excel_sync.load_from_excel(EXCEL_FILE)

def shutdown_save():
    """서버 종료 시 자동으로 Excel에 저장"""
    try:
        excel_sync.save_to_excel(mrp, EXCEL_FILE)
        logger.info("서버 종료 시 자동 저장 완료: %s", EXCEL_FILE)
    except Exception as e:
        logger.error("서버 종료 시 자동 저장 실패: %s", str(e))

atexit.register(shutdown_save)

def _backup_daemon():
    """매일 00:00에 Excel 백업 생성 후 7일 로테이션(create_backup 내장)"""
    while True:
        now = datetime.now()
        next_midnight = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        time.sleep(max((next_midnight - now).total_seconds(), 1))
        try:
            backup_path = excel_sync.create_backup(EXCEL_FILE)
            logger.info("자정 자동 백업 완료: %s", backup_path)
        except Exception as e:
            logger.error("자정 자동 백업 실패: %s", str(e))

def start_backup_scheduler():
    t = threading.Thread(target=_backup_daemon, name="backup-scheduler", daemon=True)
    t.start()
    logger.info("자정 자동 백업 스케줄러 시작 (매일 00:00)")

# Pydantic models for API requests
class ItemUpdateRequest(BaseModel):
    base_stock: float
    safety_stock: float
    target_stock: float
    lead_time: Optional[int] = 5

    @field_validator("base_stock", "safety_stock", "target_stock")
    @classmethod
    def stock_non_negative(cls, v: float) -> float:
        if v < 0:
            raise ValueError("재고 수량은 음수일 수 없습니다")
        return v

    @field_validator("lead_time")
    @classmethod
    def lead_time_positive(cls, v: Optional[int]) -> Optional[int]:
        if v is not None and v < 1:
            raise ValueError("리드타임은 1 이상이어야 합니다")
        return v

class ProductionRequest(BaseModel):
    date: str
    slip_no: str
    line: str
    fg_code: str
    prod_qty: float
    good_qty: float
    defect_qty: float
    worker: str
    note: Optional[str] = ""

    @field_validator("prod_qty", "good_qty", "defect_qty")
    @classmethod
    def qty_non_negative(cls, v: float) -> float:
        if v < 0:
            raise ValueError("수량은 음수일 수 없습니다")
        return v

    @field_validator("prod_qty")
    @classmethod
    def prod_qty_positive(cls, v: float) -> float:
        if v <= 0:
            raise ValueError("생산수량은 0보다 커야 합니다")
        return v

class TransactionRequest(BaseModel):
    date: str
    slip_no: str
    tx_type: str
    part_name: str
    qty: float
    customer: str
    note: Optional[str] = ""
    worker: Optional[str] = ""

    @field_validator("qty")
    @classmethod
    def qty_non_zero(cls, v: float) -> float:
        if v == 0:
            raise ValueError("수량은 0일 수 없습니다")
        return v

class PurchaseOrderRequest(BaseModel):
    date: str
    po_no: str
    part_name: str
    order_qty: float
    recv_qty: Optional[float] = 0.0
    exp_date: str
    customer: str
    note: Optional[str] = ""

    @field_validator("order_qty")
    @classmethod
    def order_qty_positive(cls, v: float) -> float:
        if v <= 0:
            raise ValueError("발주수량은 0보다 커야 합니다")
        return v

    @field_validator("recv_qty")
    @classmethod
    def recv_qty_non_negative(cls, v: Optional[float]) -> Optional[float]:
        if v is not None and v < 0:
            raise ValueError("입고수량은 음수일 수 없습니다")
        return v

class FGTransactionRequest(BaseModel):
    date: str
    slip_no: str
    fg_code: str
    out_qty: float
    customer: str
    note: Optional[str] = ""
    worker: Optional[str] = ""

    @field_validator("out_qty")
    @classmethod
    def out_qty_positive(cls, v: float) -> float:
        if v <= 0:
            raise ValueError("출하수량은 0보다 커야 합니다")
        return v

class PlanSimulationRequest(BaseModel):
    plans: Dict[str, float]

    @field_validator("plans")
    @classmethod
    def plans_non_negative(cls, v: Dict[str, float]) -> Dict[str, float]:
        for code, qty in v.items():
            if qty < 0:
                raise ValueError(f"{code} 계획수량은 음수일 수 없습니다")
        return v

class BOMItemRequest(BaseModel):
    fg_code: str
    fg_name: str
    part_category: str
    part_name: str
    unit_qty: float
    unit: Optional[str] = "EA"
    note: Optional[str] = ""

    @field_validator("unit_qty")
    @classmethod
    def unit_qty_positive(cls, v: float) -> float:
        if v <= 0:
            raise ValueError("투입수량은 0보다 커야 합니다")
        return v

_status_cache = {"data": None, "timestamp": 0}
CACHE_TTL = 5

def invalidate_cache():
    _status_cache["data"] = None
    _status_cache["timestamp"] = 0

@app.get("/api/status")
def get_status():
    now = time.time()
    if _status_cache["data"] and (now - _status_cache["timestamp"]) < CACHE_TTL:
        return _status_cache["data"]
    
    kpis = mrp.get_kpi_summary()
    inventory = mrp.calculate_inventory()
    capacity = mrp.calculate_capacity(inventory)
    simulation = mrp.simulate_plans()
    fg_inventory = mrp.calculate_fg_inventory()
    
    result = {
        "kpis": kpis,
        "inventory": inventory,
        "capacity": capacity,
        "simulation": simulation,
        "finished_goods": [
            {"code": fg.code, "name": fg.name, "spec": fg.spec, "unit": fg.unit, "note": fg.note}
            for fg in mrp.finished_goods.values()
        ],
        "items": [
            {"code": it.code, "name": it.name, "category": it.category, "unit": it.unit,
             "base_stock": it.base_stock, "safety_stock": it.safety_stock, "target_stock": it.target_stock,
             "lead_time": it.lead_time, "note": it.note}
            for it in mrp.items.values()
        ],
        "bom_list": [
            {"fg_code": b.fg_code, "fg_name": b.fg_name, "part_category": b.part_category,
             "part_name": b.part_name, "unit_qty": b.unit_qty, "unit": b.unit, "note": b.note}
            for b in mrp.bom_list
        ],
        "fg_inventory": fg_inventory,
        "fg_transactions": [
            {"id": ft.id, "date": ft.date, "slip_no": ft.slip_no, "fg_code": ft.fg_code,
             "fg_name": ft.fg_name, "unit": ft.unit, "out_qty": ft.out_qty,
             "customer": ft.customer, "note": ft.note, "worker": ft.worker}
            for ft in reversed(mrp.fg_transactions)
        ],
        "production_logs": [
            {"id": p.id, "date": p.date, "slip_no": p.slip_no, "line": p.line, "fg_code": p.fg_code,
             "fg_name": p.fg_name, "spec": p.spec, "prod_qty": p.prod_qty, "good_qty": p.good_qty,
             "defect_qty": p.defect_qty, "worker": p.worker, "note": p.note}
            for p in reversed(mrp.production_logs)
        ],
        "transactions": [
            {"id": t.id, "date": t.date, "slip_no": t.slip_no, "tx_type": t.tx_type,
             "part_category": t.part_category, "part_name": t.part_name, "unit": t.unit,
             "in_qty": t.in_qty, "out_qty": t.out_qty, "customer": t.customer, "note": t.note, "worker": t.worker}
            for t in reversed(mrp.transactions)
        ],
        "purchase_orders": [
            {"id": po.id, "date": po.date, "po_no": po.po_no, "part_code": po.part_code,
             "part_name": po.part_name, "unit": po.unit, "order_qty": po.order_qty,
             "recv_qty": po.recv_qty, "pending_qty": po.pending_qty, "exp_date": po.exp_date,
             "customer": po.customer, "status": po.status, "note": po.note}
            for po in reversed(mrp.purchase_orders)
        ]
    }
    
    _status_cache["data"] = result
    _status_cache["timestamp"] = now
    return result

# --- Item Master Edit ---
@app.put("/api/item/{code}")
def update_item_endpoint(code: str, req: ItemUpdateRequest):
    try:
        updated = mrp.update_item(
            code=code, base_stock=req.base_stock, safety_stock=req.safety_stock,
            target_stock=req.target_stock, lead_time=req.lead_time or 5
        )
        logger.info("품목 수정 완료: code=%s, name=%s", code, updated.name)
        invalidate_cache()
        return {"success": True, "message": f"품목 [{updated.name}] 재고 기준정보가 수정되었습니다."}
    except Exception as e:
        logger.error("품목 수정 실패: code=%s, error=%s", code, str(e))
        raise HTTPException(status_code=400, detail=str(e))

# --- Production Logs CRUD ---
@app.post("/api/production")
def add_production(req: ProductionRequest):
    try:
        log = mrp.add_production_log(
            date=req.date, slip_no=req.slip_no, line=req.line, fg_code=req.fg_code,
            prod_qty=req.prod_qty, good_qty=req.good_qty, defect_qty=req.defect_qty,
            worker=req.worker, note=req.note or ""
        )
        logger.info("생산 실적 등록: id=%d, fg_code=%s, prod_qty=%d", log.id, req.fg_code, req.prod_qty)
        invalidate_cache()
        return {"success": True, "message": "생산 실적이 성공적으로 등록되었습니다 (BOM 자재 자동 차감 완료)", "id": log.id}
    except Exception as e:
        logger.error("생산 실적 등록 실패: fg_code=%s, error=%s", req.fg_code, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.put("/api/production/{log_id}")
def update_production(log_id: int, req: ProductionRequest):
    try:
        log = mrp.update_production_log(
            log_id=log_id, date=req.date, slip_no=req.slip_no, line=req.line,
            fg_code=req.fg_code, prod_qty=req.prod_qty, good_qty=req.good_qty,
            defect_qty=req.defect_qty, worker=req.worker, note=req.note or ""
        )
        logger.info("생산 실적 수정: id=%d, fg_code=%s, prod_qty=%d", log_id, req.fg_code, req.prod_qty)
        invalidate_cache()
        return {"success": True, "message": f"생산 실적 [전표 {log.slip_no}] 수정 완료 (자재 자동 재계산)", "id": log.id}
    except Exception as e:
        logger.error("생산 실적 수정 실패: id=%d, error=%s", log_id, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.delete("/api/production/{log_id}")
def delete_production(log_id: int):
    try:
        mrp.delete_production_log(log_id)
        logger.info("생산 실적 삭제: id=%d", log_id)
        invalidate_cache()
        return {"success": True, "message": f"생산 실적 ID {log_id} 삭제 완료 (차감 자재 원복 완료)"}
    except Exception as e:
        logger.error("생산 실적 삭제 실패: id=%d, error=%s", log_id, str(e))
        raise HTTPException(status_code=400, detail=str(e))

# --- Transactions CRUD ---
@app.post("/api/transaction")
def add_transaction(req: TransactionRequest):
    try:
        tx = mrp.add_transaction(
            date=req.date, slip_no=req.slip_no, tx_type=req.tx_type,
            part_name=req.part_name, qty=req.qty, customer=req.customer,
            note=req.note or "", worker=req.worker or ""
        )
        logger.info("입출고 등록: id=%d, tx_type=%s, part=%s, qty=%d", tx.id, req.tx_type, req.part_name, req.qty)
        invalidate_cache()
        return {"success": True, "message": f"자재 {req.tx_type} 등록 완료", "id": tx.id}
    except Exception as e:
        logger.error("입출고 등록 실패: tx_type=%s, part=%s, error=%s", req.tx_type, req.part_name, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.put("/api/transaction/{tx_id}")
def update_transaction_endpoint(tx_id: int, req: TransactionRequest):
    try:
        tx = mrp.update_transaction(
            tx_id=tx_id, date=req.date, slip_no=req.slip_no, tx_type=req.tx_type,
            part_name=req.part_name, qty=req.qty, customer=req.customer,
            note=req.note or "", worker=req.worker or ""
        )
        logger.info("입출고 수정: id=%d, tx_type=%s, part=%s, qty=%d", tx_id, req.tx_type, req.part_name, req.qty)
        invalidate_cache()
        return {"success": True, "message": f"자재 {tx.tx_type} 전표 [ID {tx.id}] 수정 완료", "id": tx.id}
    except Exception as e:
        logger.error("입출고 수정 실패: id=%d, error=%s", tx_id, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.delete("/api/transaction/{tx_id}")
def delete_transaction_endpoint(tx_id: int):
    try:
        mrp.delete_transaction(tx_id)
        logger.info("입출고 삭제: id=%d", tx_id)
        invalidate_cache()
        return {"success": True, "message": f"입출고 전표 ID {tx_id} 삭제 완료"}
    except Exception as e:
        logger.error("입출고 삭제 실패: id=%d, error=%s", tx_id, str(e))
        raise HTTPException(status_code=400, detail=str(e))

# --- Finished Goods (완제품 출하/판매) CRUD ---
@app.post("/api/fgtransaction")
def add_fg_transaction(req: FGTransactionRequest):
    try:
        fg_tx = mrp.add_fg_transaction(
            date=req.date, slip_no=req.slip_no, fg_code=req.fg_code,
            out_qty=req.out_qty, customer=req.customer,
            note=req.note or "", worker=req.worker or ""
        )
        logger.info("완제품 출하 등록: id=%d, fg=%s, qty=%d", fg_tx.id, req.fg_code, req.out_qty)
        invalidate_cache()
        return {"success": True, "message": f"완제품 {req.fg_code} 출하 등록 완료 (완제품 재고 자동 차감)", "id": fg_tx.id}
    except Exception as e:
        logger.error("완제품 출하 등록 실패: fg_code=%s, error=%s", req.fg_code, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.put("/api/fgtransaction/{fg_tx_id}")
def update_fg_transaction(fg_tx_id: int, req: FGTransactionRequest):
    try:
        fg_tx = mrp.update_fg_transaction(
            fg_tx_id=fg_tx_id, date=req.date, slip_no=req.slip_no, fg_code=req.fg_code,
            out_qty=req.out_qty, customer=req.customer,
            note=req.note or "", worker=req.worker or ""
        )
        logger.info("완제품 출하 수정: id=%d, fg=%s, qty=%d", fg_tx_id, req.fg_code, req.out_qty)
        invalidate_cache()
        return {"success": True, "message": f"완제품 출하 전표 [ID {fg_tx.id}] 수정 완료 (완제품 재고 자동 재계산)", "id": fg_tx.id}
    except Exception as e:
        logger.error("완제품 출하 수정 실패: id=%d, error=%s", fg_tx_id, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.delete("/api/fgtransaction/{fg_tx_id}")
def delete_fg_transaction(fg_tx_id: int):
    try:
        mrp.delete_fg_transaction(fg_tx_id)
        logger.info("완제품 출하 삭제: id=%d", fg_tx_id)
        invalidate_cache()
        return {"success": True, "message": f"완제품 출하 전표 ID {fg_tx_id} 삭제 완료 (완제품 재고 원복)"}
    except Exception as e:
        logger.error("완제품 출하 삭제 실패: id=%d, error=%s", fg_tx_id, str(e))
        raise HTTPException(status_code=400, detail=str(e))

# --- Purchase Orders CRUD ---
@app.post("/api/po")
def add_purchase_order(req: PurchaseOrderRequest):
    try:
        po = mrp.add_purchase_order(
            date=req.date, po_no=req.po_no, part_name=req.part_name,
            order_qty=req.order_qty, exp_date=req.exp_date,
            customer=req.customer, recv_qty=req.recv_qty or 0.0,
            note=req.note or ""
        )
        logger.info("발주 등록: id=%d, part=%s, order_qty=%d", po.id, req.part_name, req.order_qty)
        invalidate_cache()
        return {"success": True, "message": "발주 등록 완료 (미입고량 자동 반영)", "id": po.id}
    except Exception as e:
        logger.error("발주 등록 실패: part=%s, error=%s", req.part_name, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.put("/api/po/{po_id}")
def update_purchase_order_endpoint(po_id: int, req: PurchaseOrderRequest):
    try:
        po = mrp.update_purchase_order(
            po_id=po_id, date=req.date, po_no=req.po_no, part_name=req.part_name,
            order_qty=req.order_qty, recv_qty=req.recv_qty or 0.0,
            exp_date=req.exp_date, customer=req.customer, note=req.note or ""
        )
        logger.info("발주 수정: id=%d, part=%s, status=%s", po_id, req.part_name, po.status)
        invalidate_cache()
        return {"success": True, "message": f"발주 전표 [번호 {po.po_no}] 수정 완료 (상태: {po.status})", "id": po.id}
    except Exception as e:
        logger.error("발주 수정 실패: id=%d, error=%s", po_id, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.delete("/api/po/{po_id}")
def delete_purchase_order_endpoint(po_id: int):
    try:
        mrp.delete_purchase_order(po_id)
        logger.info("발주 삭제: id=%d", po_id)
        invalidate_cache()
        return {"success": True, "message": f"발주 전표 ID {po_id} 삭제 완료"}
    except Exception as e:
        logger.error("발주 삭제 실패: id=%d, error=%s", po_id, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.get("/api/po/{po_id}/print")
def print_purchase_order(po_id: int):
    po = next((p for p in mrp.purchase_orders if p.id == po_id), None)
    if not po:
        raise HTTPException(status_code=404, detail=f"발주 ID {po_id}를 찾을 수 없습니다.")
    
    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>발주서 - {po.po_no}</title>
    <style>
        @media print {{ body {{ margin: 0; }} .no-print {{ display: none; }} }}
        body {{ font-family: 'Malgun Gothic', sans-serif; padding: 40px; }}
        .header {{ text-align: center; border-bottom: 3px solid #333; padding-bottom: 20px; margin-bottom: 30px; }}
        .header h1 {{ font-size: 28px; margin: 0; }}
        .info-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }}
        .info-box {{ border: 1px solid #ddd; padding: 15px; border-radius: 8px; }}
        .info-box h3 {{ margin: 0 0 10px 0; color: #555; font-size: 14px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        th, td {{ border: 1px solid #333; padding: 12px; text-align: center; }}
        th {{ background-color: #f5f5f5; font-weight: bold; }}
        .footer {{ margin-top: 50px; display: grid; grid-template-columns: 1fr 1fr; gap: 100px; }}
        .sign-box {{ border-top: 2px solid #333; padding-top: 10px; text-align: center; }}
        .btn-print {{ background: #3b82f6; color: white; padding: 10px 30px; border: none; border-radius: 8px; cursor: pointer; font-size: 16px; margin-top: 20px; }}
    </style>
</head>
<body>
    <div class="no-print" style="text-align: right;">
        <button class="btn-print" onclick="window.print()"> 발주서 인쇄 </button>
    </div>
    <div class="header">
        <h1>원 부 자재 발 주 서</h1>
    </div>
    <div class="info-grid">
        <div class="info-box">
            <h3>■ 발주 정보</h3>
            <p><strong>발주번호:</strong> {po.po_no}</p>
            <p><strong>발주일:</strong> {po.date}</p>
            <p><strong>예정입고일:</strong> {po.exp_date}</p>
        </div>
        <div class="info-box">
            <h3>■ 거래처 정보</h3>
            <p><strong>거래처:</strong> {po.customer}</p>
            <p><strong>상태:</strong> {po.status}</p>
        </div>
    </div>
    <table>
        <thead>
            <tr>
                <th>부품코드</th>
                <th>부품명</th>
                <th>단위</th>
                <th>발주수량</th>
                <th>입고수량</th>
                <th>미입고수량</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>{po.part_code}</td>
                <td>{po.part_name}</td>
                <td>{po.unit}</td>
                <td>{po.order_qty:,.0f}</td>
                <td>{po.recv_qty:,.0f}</td>
                <td>{po.pending_qty:,.0f}</td>
            </tr>
        </tbody>
    </table>
    <div class="footer">
        <div class="sign-box">
            <p>담당자</p>
        </div>
        <div class="sign-box">
            <p>승인자 (인)</p>
        </div>
    </div>
</body>
</html>"""
    from fastapi.responses import HTMLResponse
    return HTMLResponse(content=html)

class BatchPrintPORequest(BaseModel):
    po_ids: list[int] = []

@app.post("/api/po/batch_print")
def batch_print_purchase_orders(req: BatchPrintPORequest):
    po_ids = req.po_ids or []
    pos = [p for p in mrp.purchase_orders if p.id in po_ids]
    if not pos:
        raise HTTPException(status_code=404, detail="선택된 발주가 없습니다.")
    rows = []
    for po in sorted(pos, key=lambda x: x.id):
        rows.append(f"""
        <tr>
          <td>{po.po_no}</td>
          <td>{po.date}</td>
          <td>{po.part_code}</td>
          <td>{po.part_name}</td>
          <td>{po.unit}</td>
          <td>{po.order_qty:,.0f}</td>
          <td>{po.recv_qty:,.0f}</td>
          <td>{po.pending_qty:,.0f}</td>
          <td>{po.exp_date}</td>
          <td>{po.customer}</td>
          <td>{po.status}</td>
        </tr>
        """)
    html = f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8"><title>발주서 일괄 출력</title>
<style>
@media print {{ body {{ margin:0; }} .no-print {{ display:none; }} }}
body {{ font-family:'Malgun Gothic',sans-serif; padding:30px; }}
.header {{ text-align:center; border-bottom:3px solid #333; padding-bottom:15px; margin-bottom:20px; }}
.header h1 {{ font-size:24px; margin:0; }}
table {{ width:100%; border-collapse:collapse; margin-top:10px; font-size:12px; }}
th,td {{ border:1px solid #333; padding:8px; text-align:center; }}
th {{ background:#f5f5f5; font-weight:bold; }}
.btn-print {{ background:#3b82f6; color:white; padding:10px 30px; border:none; border-radius:8px; cursor:pointer; font-size:16px; margin-top:20px; }}
</style></head><body>
<div class="no-print" style="text-align:right;"><button class="btn-print" onclick="window.print()"> 일괄 인쇄 </button></div>
<div class="header"><h1>원부자재 발주 및 미입고 관리 대장 일괄 출력</h1></div>
<table>
<thead><tr>
<th>발주번호</th><th>발주일</th><th>부품코드</th><th>부품명</th><th>단위</th><th>발주수량</th><th>입고수량</th><th>미입고수량</th><th>예정입고일</th><th>거래처</th><th>상태</th>
</tr></thead>
<tbody>
{''.join(rows)}
</tbody>
</table>
</body></html>"""
    from fastapi.responses import HTMLResponse
    return HTMLResponse(content=html)

class BatchPORequest(BaseModel):
    customer: str = ""
    exp_date: str = ""
    note: str = ""

@app.post("/api/po/batch")
def generate_batch_purchase_orders(req: BatchPORequest):
    try:
        inv_rows = mrp.calculate_inventory()
        new_pos = []
        
        for item in inv_rows:
            if item["order_needed"] > 0:
                po_no = f"BATCH-{item['code']}"
                po = mrp.add_purchase_order(
                    date=datetime.now().strftime("%Y-%m-%d"),
                    po_no=po_no,
                    part_name=item["name"],
                    order_qty=item["order_needed"],
                    exp_date=req.exp_date or datetime.now().strftime("%Y-%m-%d"),
                    customer=req.customer or "자동발주",
                    recv_qty=0.0,
                    note=req.note or f"안전재고 기반 자동 발주 (필요량: {item['order_needed']:.0f} EA)"
                )
                new_pos.append({"id": po.id, "part_name": item["name"], "order_qty": item["order_needed"]})
                logger.info("일괄 발주 생성: id=%d, part=%s, qty=%d", po.id, item["name"], item["order_needed"])
        
        invalidate_cache()
        return {"success": True, "message": f"{len(new_pos)}건의 발주서가 자동 생성되었습니다.", "orders": new_pos}
    except Exception as e:
        logger.error("일괄 발주 생성 실패: %s", str(e))
        raise HTTPException(status_code=400, detail=str(e))

# --- Plan & BOM ---
@app.post("/api/plan")
def update_plan(req: PlanSimulationRequest):
    sim = mrp.simulate_plans(req.plans)
    return {"success": True, "simulation": sim}

@app.post("/api/bom")
def add_bom_item(req: BOMItemRequest):
    b = BOMItem(
        fg_code=req.fg_code, fg_name=req.fg_name, part_category=req.part_category,
        part_name=req.part_name, unit_qty=req.unit_qty, unit=req.unit or "EA", note=req.note or ""
    )
    mrp.add_bom_item(b)
    invalidate_cache()
    return {"success": True, "message": f"BOM 부품 등록 완료: {req.fg_code} -> {req.part_name} ({req.unit_qty} EA)"}

@app.put("/api/bom/{bom_index}")
def update_bom_item(bom_index: int, req: BOMItemRequest):
    try:
        if bom_index < 0 or bom_index >= len(mrp.bom_list):
            raise HTTPException(status_code=404, detail=f"BOM 인덱스 {bom_index}를 찾을 수 없습니다.")
        
        old = mrp.bom_list[bom_index]
        mrp.bom_list[bom_index] = BOMItem(
            fg_code=req.fg_code, fg_name=req.fg_name, part_category=req.part_category,
            part_name=req.part_name, unit_qty=req.unit_qty, unit=req.unit or "EA", note=req.note or ""
        )
        logger.info("BOM 수정: index=%d, fg=%s -> part=%s", bom_index, req.fg_code, req.part_name)
        invalidate_cache()
        return {"success": True, "message": f"BOM 수정 완료: {old.fg_code} -> {req.part_name}"}
    except HTTPException:
        raise
    except Exception as e:
        logger.error("BOM 수정 실패: index=%d, error=%s", bom_index, str(e))
        raise HTTPException(status_code=400, detail=str(e))

@app.delete("/api/bom/{bom_index}")
def delete_bom_item(bom_index: int):
    try:
        if bom_index < 0 or bom_index >= len(mrp.bom_list):
            raise HTTPException(status_code=404, detail=f"BOM 인덱스 {bom_index}를 찾을 수 없습니다.")
        
        removed = mrp.bom_list.pop(bom_index)
        logger.info("BOM 삭제: index=%d, fg=%s, part=%s", bom_index, removed.fg_code, removed.part_name)
        invalidate_cache()
        return {"success": True, "message": f"BOM 삭제 완료: {removed.fg_code} -> {removed.part_name}"}
    except HTTPException:
        raise
    except Exception as e:
        logger.error("BOM 삭제 실패: index=%d, error=%s", bom_index, str(e))
        raise HTTPException(status_code=400, detail=str(e))

# --- Excel Sync ---
@app.post("/api/sync/save")
def sync_save():
    try:
        excel_sync.save_to_excel(mrp, EXCEL_FILE)
        logger.info("엑셀 저장 완료: %s", EXCEL_FILE)
        invalidate_cache()
        return {"success": True, "message": f"성공적으로 엑셀 파일({os.path.basename(EXCEL_FILE)})에 저장되었습니다."}
    except Exception as e:
        logger.error("엑셀 저장 실패: %s, error=%s", EXCEL_FILE, str(e))
        raise HTTPException(status_code=500, detail=f"엑셀 저장 오류: {str(e)}")

@app.post("/api/sync/load")
def sync_load():
    global mrp
    try:
        mrp = excel_sync.load_from_excel(EXCEL_FILE)
        logger.info("엑셀 로드 완료: %s", EXCEL_FILE)
        invalidate_cache()
        return {"success": True, "message": f"엑셀 파일({os.path.basename(EXCEL_FILE)})로부터 데이터를 다시 로드했습니다."}
    except Exception as e:
        logger.error("엑셀 로드 실패: %s, error=%s", EXCEL_FILE, str(e))
        raise HTTPException(status_code=500, detail=f"엑셀 로드 오류: {str(e)}")

# Mount static folder
static_dir = os.path.join(base_path, "static")
os.makedirs(static_dir, exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")

@app.get("/")
def read_index():
    index_path = os.path.join(static_dir, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return JSONResponse({"status": "MRP Backend API running. Please place index.html in static/"})

@app.get("/favicon.ico")
def favicon():
    from fastapi.responses import Response
    return Response(status_code=204)

def open_browser():
    webbrowser.open("http://127.0.0.1:8000")

if __name__ == "__main__":
    start_backup_scheduler()
    threading.Timer(1.2, open_browser).start()
    logger.info("="*60)
    logger.info("  KMTECH MRP Real-Time System Starting...")
    logger.info("  URL: http://127.0.0.1:8000")
    logger.info("="*60)
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
