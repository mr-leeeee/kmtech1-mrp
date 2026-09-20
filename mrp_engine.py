from dataclasses import dataclass, field, asdict
from typing import List, Dict, Optional, Any
import datetime
import math

def validate_positive_qty(value: float, field_name: str = "수량") -> float:
    q = float(value)
    if q <= 0:
        raise ValueError(f"{field_name}은(는) 0보다 커야 합니다: {q}")
    return q

def validate_non_negative_qty(value: float, field_name: str = "수량") -> float:
    q = float(value)
    if q < 0:
        raise ValueError(f"{field_name}은(는) 음수일 수 없습니다: {q}")
    return q

def validate_fg_code_exists(mrp: "MRPSystem", fg_code: str) -> str:
    if fg_code not in mrp.finished_goods:
        raise KeyError(f"완제품 코드를 찾을 수 없습니다: {fg_code}")
    return fg_code

@dataclass
class Item:
    code: str
    name: str
    category: str
    unit: str
    base_stock: float
    safety_stock: float
    target_stock: float
    lead_time: int = 5
    note: str = ""

@dataclass
class FinishedGood:
    code: str
    name: str
    spec: str
    unit: str = "Set"
    note: str = ""

@dataclass
class BOMItem:
    fg_code: str
    fg_name: str
    part_category: str
    part_name: str
    unit_qty: float
    unit: str = "EA"
    note: str = ""

@dataclass
class ProductionLog:
    id: int
    date: str
    slip_no: str
    line: str
    fg_code: str
    fg_name: str
    spec: str
    prod_qty: float
    good_qty: float
    defect_qty: float
    worker: str
    note: str = ""

@dataclass
class Transaction:
    id: int
    date: str
    slip_no: str
    tx_type: str # '입고' or '출고'
    part_category: str
    part_name: str
    unit: str
    in_qty: float
    out_qty: float
    customer: str
    note: str = ""
    worker: str = ""

@dataclass
class PurchaseOrder:
    id: int
    date: str
    po_no: str
    part_code: str
    part_name: str
    unit: str
    order_qty: float
    recv_qty: float
    pending_qty: float
    exp_date: str
    customer: str
    status: str # '미입고', '부분입고', '입고완료'
    note: str = ""

@dataclass
class FGTransaction:
    """완제품 출하/판매 거래 (완제품 재고 차감용)"""
    id: int
    date: str
    slip_no: str
    fg_code: str
    fg_name: str
    unit: str
    out_qty: float
    customer: str
    note: str = ""
    worker: str = ""

class MRPSystem:
    def __init__(self):
        self.items: Dict[str, Item] = {}           # key: part_name
        self.items_by_code: Dict[str, Item] = {}   # key: part_code
        self.finished_goods: Dict[str, FinishedGood] = {} # key: fg_code
        self.bom_list: List[BOMItem] = []
        self.production_logs: List[ProductionLog] = []
        self.transactions: List[Transaction] = []
        self.purchase_orders: List[PurchaseOrder] = []
        self.fg_transactions: List[FGTransaction] = []  # 완제품 출하/판매
        self.production_plans: Dict[str, float] = {} # key: fg_code -> qty
        self._next_prod_id = 1
        self._next_tx_id = 1
        self._next_po_id = 1
        self._next_fg_tx_id = 1

    def add_item(self, item: Item):
        self.items[item.name] = item
        self.items_by_code[item.code] = item

    def update_item(self, code: str, base_stock: float, safety_stock: float,
                    target_stock: float, lead_time: int = 5) -> Item:
        item = self.items_by_code.get(code)
        if not item:
            raise KeyError(f"품목코드를 찾을 수 없습니다: {code}")
        item.base_stock = float(base_stock)
        item.safety_stock = float(safety_stock)
        item.target_stock = float(target_stock)
        item.lead_time = int(lead_time)
        return item

    def add_finished_good(self, fg: FinishedGood):
        self.finished_goods[fg.code] = fg

    def add_bom_item(self, bom: BOMItem):
        self.bom_list.append(bom)

    def add_production_log(self, date: str, slip_no: str, line: str, fg_code: str,
                           prod_qty: float, good_qty: float, defect_qty: float,
                           worker: str, note: str = "") -> ProductionLog:
        fg = self.finished_goods.get(fg_code)
        fg_name = fg.name if fg else "미등록 완제품"
        spec = fg.spec if fg else ""
        
        # Validation
        if prod_qty != (good_qty + defect_qty):
            raise ValueError(f"수량 불일치 오류: 생산수량({prod_qty}) != 양품({good_qty}) + 불량({defect_qty})")
        if prod_qty < 0 or good_qty < 0 or defect_qty < 0:
            raise ValueError("음수 수량 입력 오류")

        new_id = self._next_prod_id
        self._next_prod_id += 1
        log = ProductionLog(
            id=new_id, date=date, slip_no=slip_no, line=line,
            fg_code=fg_code, fg_name=fg_name, spec=spec,
            prod_qty=prod_qty, good_qty=good_qty, defect_qty=defect_qty,
            worker=worker, note=note
        )
        self.production_logs.append(log)
        return log

    def update_production_log(self, log_id: int, date: str, slip_no: str, line: str,
                              fg_code: str, prod_qty: float, good_qty: float,
                              defect_qty: float, worker: str, note: str = "") -> ProductionLog:
        for log in self.production_logs:
            if log.id == log_id:
                if prod_qty != (good_qty + defect_qty):
                    raise ValueError(f"수량 불일치 오류: 생산수량({prod_qty}) != 양품({good_qty}) + 불량({defect_qty})")
                if prod_qty < 0 or good_qty < 0 or defect_qty < 0:
                    raise ValueError("음수 수량 입력 오류")
                
                fg = self.finished_goods.get(fg_code)
                log.date = date
                log.slip_no = slip_no
                log.line = line
                log.fg_code = fg_code
                log.fg_name = fg.name if fg else "미등록 완제품"
                log.spec = fg.spec if fg else ""
                log.prod_qty = prod_qty
                log.good_qty = good_qty
                log.defect_qty = defect_qty
                log.worker = worker
                log.note = note
                return log
        raise KeyError(f"생산실적 ID {log_id}를 찾을 수 없습니다.")

    def delete_production_log(self, log_id: int):
        orig_len = len(self.production_logs)
        self.production_logs = [log for log in self.production_logs if log.id != log_id]
        if len(self.production_logs) == orig_len:
            raise KeyError(f"생산실적 ID {log_id}를 찾을 수 없습니다.")

    def add_transaction(self, date: str, slip_no: str, tx_type: str, part_name: str,
                        qty: float, customer: str, note: str = "", worker: str = "") -> Transaction:
        item = self.items.get(part_name)
        part_category = item.category if item else "미등록 품목"
        unit = item.unit if item else "EA"
        
        if qty < 0:
            raise ValueError("음수 수량 입력 오류")

        in_qty = qty if tx_type == "입고" else 0.0
        out_qty = qty if tx_type == "출고" else 0.0

        new_id = self._next_tx_id
        self._next_tx_id += 1
        tx = Transaction(
            id=new_id, date=date, slip_no=slip_no, tx_type=tx_type,
            part_category=part_category, part_name=part_name, unit=unit,
            in_qty=in_qty, out_qty=out_qty, customer=customer, note=note, worker=worker
        )
        self.transactions.append(tx)
        return tx

    def update_transaction(self, tx_id: int, date: str, slip_no: str, tx_type: str,
                           part_name: str, qty: float, customer: str, note: str = "",
                           worker: str = "") -> Transaction:
        for tx in self.transactions:
            if tx.id == tx_id:
                if qty < 0:
                    raise ValueError("음수 수량 입력 오류")
                item = self.items.get(part_name)
                tx.date = date
                tx.slip_no = slip_no
                tx.tx_type = tx_type
                tx.part_name = part_name
                tx.part_category = item.category if item else "미등록 품목"
                tx.unit = item.unit if item else "EA"
                tx.in_qty = qty if tx_type == "입고" else 0.0
                tx.out_qty = qty if tx_type == "출고" else 0.0
                tx.customer = customer
                tx.note = note
                tx.worker = worker
                return tx
        raise KeyError(f"입출고 ID {tx_id}를 찾을 수 없습니다.")

    def delete_transaction(self, tx_id: int):
        orig_len = len(self.transactions)
        self.transactions = [tx for tx in self.transactions if tx.id != tx_id]
        if len(self.transactions) == orig_len:
            raise KeyError(f"입출고 ID {tx_id}를 찾을 수 없습니다.")

    # --- 완제품 출하(판매) 거래 관리 ---
    def add_fg_transaction(self, date: str, slip_no: str, fg_code: str, out_qty: float,
                           customer: str, note: str = "", worker: str = "") -> FGTransaction:
        fg = self.finished_goods.get(fg_code)
        if not fg:
            raise KeyError(f"완제품 코드를 찾을 수 없습니다: {fg_code}")

        if out_qty <= 0:
            raise ValueError("출하수량은 0보다 커야 합니다.")

        new_id = self._next_fg_tx_id
        self._next_fg_tx_id += 1
        fg_tx = FGTransaction(
            id=new_id, date=date, slip_no=slip_no, fg_code=fg_code,
            fg_name=fg.name, unit=fg.unit, out_qty=out_qty,
            customer=customer, note=note, worker=worker
        )
        self.fg_transactions.append(fg_tx)
        return fg_tx

    def update_fg_transaction(self, fg_tx_id: int, date: str, slip_no: str, fg_code: str,
                              out_qty: float, customer: str, note: str = "",
                              worker: str = "") -> FGTransaction:
        fg = self.finished_goods.get(fg_code)
        if not fg:
            raise KeyError(f"완제품 코드를 찾을 수 없습니다: {fg_code}")
        if out_qty <= 0:
            raise ValueError("출하수량은 0보다 커야 합니다.")

        for fg_tx in self.fg_transactions:
            if fg_tx.id == fg_tx_id:
                fg_tx.date = date
                fg_tx.slip_no = slip_no
                fg_tx.fg_code = fg_code
                fg_tx.fg_name = fg.name
                fg_tx.unit = fg.unit
                fg_tx.out_qty = out_qty
                fg_tx.customer = customer
                fg_tx.note = note
                fg_tx.worker = worker
                return fg_tx
        raise KeyError(f"완제품 출하 ID {fg_tx_id}를 찾을 수 없습니다.")

    def delete_fg_transaction(self, fg_tx_id: int):
        orig_len = len(self.fg_transactions)
        self.fg_transactions = [t for t in self.fg_transactions if t.id != fg_tx_id]
        if len(self.fg_transactions) == orig_len:
            raise KeyError(f"완제품 출하 ID {fg_tx_id}를 찾을 수 없습니다.")

    def calculate_fg_inventory(self) -> List[Dict[str, Any]]:
        """
        완제품 현재고 계산: 누적 양품 생산량(good_qty) - 누적 출하량(out_qty)
        상태: 정상 / 주의 / 부족
        """
        # 1. 완제품별 누적 양품 생산량
        cum_good_by_fg: Dict[str, float] = {}
        for log in self.production_logs:
            cum_good_by_fg[log.fg_code] = cum_good_by_fg.get(log.fg_code, 0.0) + log.good_qty

        # 2. 완제품별 누적 출하량
        cum_out_by_fg: Dict[str, float] = {}
        for t in self.fg_transactions:
            cum_out_by_fg[t.fg_code] = cum_out_by_fg.get(t.fg_code, 0.0) + t.out_qty

        # 3. 완제품별 현재고 및 상태 산출
        fg_inv_rows = []
        for fg in self.finished_goods.values():
            produced = cum_good_by_fg.get(fg.code, 0.0)
            shipped = cum_out_by_fg.get(fg.code, 0.0)
            current = produced - shipped

            if current < 0:
                status = "부족"
            elif current == 0:
                status = "주문 대기 (재고 0)"
            else:
                status = "정상"

            fg_inv_rows.append({
                "fg_code": fg.code,
                "fg_name": fg.name,
                "spec": fg.spec,
                "unit": fg.unit,
                "produced_qty": produced,
                "shipped_qty": shipped,
                "current_stock": current,
                "status": status
            })

        return fg_inv_rows

    def add_purchase_order(self, date: str, po_no: str, part_name: str,
                           order_qty: float, exp_date: str, customer: str,
                           recv_qty: float = 0.0, note: str = "") -> PurchaseOrder:
        item = self.items.get(part_name)
        part_code = item.code if item else ""
        unit = item.unit if item else "EA"
        
        if order_qty < 0 or recv_qty < 0:
            raise ValueError("음수 수량 입력 오류")

        pending_qty = max(order_qty - recv_qty, 0.0)
        status = "미입고"
        if order_qty > 0:
            if recv_qty >= order_qty:
                status = "입고완료"
            elif recv_qty > 0:
                status = "부분입고"

        new_id = self._next_po_id
        self._next_po_id += 1
        po = PurchaseOrder(
            id=new_id, date=date, po_no=po_no, part_code=part_code,
            part_name=part_name, unit=unit, order_qty=order_qty,
            recv_qty=recv_qty, pending_qty=pending_qty, exp_date=exp_date,
            customer=customer, status=status, note=note
        )
        self.purchase_orders.append(po)
        return po

    def update_purchase_order(self, po_id: int, date: str, po_no: str, part_name: str,
                              order_qty: float, recv_qty: float, exp_date: str,
                              customer: str, note: str = "") -> PurchaseOrder:
        for po in self.purchase_orders:
            if po.id == po_id:
                if order_qty < 0 or recv_qty < 0:
                    raise ValueError("음수 수량 입력 오류")
                item = self.items.get(part_name)
                pending_qty = max(order_qty - recv_qty, 0.0)
                status = "미입고"
                if order_qty > 0:
                    if recv_qty >= order_qty:
                        status = "입고완료"
                    elif recv_qty > 0:
                        status = "부분입고"

                po.date = date
                po.po_no = po_no
                po.part_name = part_name
                po.part_code = item.code if item else ""
                po.unit = item.unit if item else "EA"
                po.order_qty = order_qty
                po.recv_qty = recv_qty
                po.pending_qty = pending_qty
                po.exp_date = exp_date
                po.customer = customer
                po.status = status
                po.note = note
                return po
        raise KeyError(f"발주 ID {po_id}를 찾을 수 없습니다.")

    def delete_purchase_order(self, po_id: int):
        orig_len = len(self.purchase_orders)
        self.purchase_orders = [po for po in self.purchase_orders if po.id != po_id]
        if len(self.purchase_orders) == orig_len:
            raise KeyError(f"발주 ID {po_id}를 찾을 수 없습니다.")

    def calculate_inventory(self) -> List[Dict[str, Any]]:
        """
        원부자재 재고 수불 및 상태 계산
        """
        # 1. 완제품별 누적 생산량 집계
        cum_prod_by_fg: Dict[str, float] = {}
        for log in self.production_logs:
            cum_prod_by_fg[log.fg_code] = cum_prod_by_fg.get(log.fg_code, 0.0) + log.prod_qty

        # 2. BOM 전개: 부품별 총 생산소요량 계산
        prod_deduct_by_part: Dict[str, float] = {}
        for b in self.bom_list:
            fg_prod = cum_prod_by_fg.get(b.fg_code, 0.0)
            req = b.unit_qty * fg_prod
            prod_deduct_by_part[b.part_name] = prod_deduct_by_part.get(b.part_name, 0.0) + req

        # 3. 입고 및 출고 누계
        in_qty_by_part: Dict[str, float] = {}
        out_qty_by_part: Dict[str, float] = {}
        for tx in self.transactions:
            if tx.tx_type == "입고":
                in_qty_by_part[tx.part_name] = in_qty_by_part.get(tx.part_name, 0.0) + tx.in_qty
            elif tx.tx_type == "출고":
                out_qty_by_part[tx.part_name] = out_qty_by_part.get(tx.part_name, 0.0) + tx.out_qty

        # 4. 미입고 발주량 집계
        pending_po_by_part: Dict[str, float] = {}
        for po in self.purchase_orders:
            pending_po_by_part[po.part_name] = pending_po_by_part.get(po.part_name, 0.0) + po.pending_qty

        # 5. 각 부품별 최종 수불 및 상태 산출
        inventory_rows = []
        for item in self.items.values():
            base = item.base_stock
            in_acc = in_qty_by_part.get(item.name, 0.0)
            out_acc = out_qty_by_part.get(item.name, 0.0)
            deduct = prod_deduct_by_part.get(item.name, 0.0)
            current = base + in_acc - out_acc - deduct
            pending = pending_po_by_part.get(item.name, 0.0)
            available = current + pending
            safety = item.safety_stock
            target = item.target_stock

            if current < 0:
                status = "부족"
            elif current < safety:
                status = "주의"
            else:
                status = "정상"

            order_needed = max(target - current - pending, 0.0)

            inventory_rows.append({
                "code": item.code,
                "name": item.name,
                "category": item.category,
                "unit": item.unit,
                "base_stock": base,
                "in_acc": in_acc,
                "out_acc": out_acc,
                "prod_deduct": deduct,
                "current_stock": current,
                "pending_po": pending,
                "available_stock": available,
                "safety_stock": safety,
                "target_stock": target,
                "status": status,
                "order_needed": order_needed,
                "lead_time": item.lead_time
            })

        return inventory_rows

    def calculate_capacity(self, inventory_rows: Optional[List[Dict[str, Any]]] = None) -> List[Dict[str, Any]]:
        """
        완제품별 실시간 생산 가용량 및 병목부품 분석
        """
        if inventory_rows is None:
            inventory_rows = self.calculate_inventory()
        inv_map = {row["name"]: row for row in inventory_rows}
        
        # 완제품별 구성 부품 매핑
        bom_by_fg: Dict[str, List[BOMItem]] = {}
        for b in self.bom_list:
            bom_by_fg.setdefault(b.fg_code, []).append(b)

        capacity_rows = []
        for fg in self.finished_goods.values():
            bom_items = bom_by_fg.get(fg.code, [])
            if not bom_items:
                capacity_rows.append({
                    "fg_code": fg.code,
                    "fg_name": fg.name,
                    "max_capacity": 0,
                    "bottleneck_part": "BOM 미등록",
                    "status": "BOM 없음",
                    "potential_capacity": 0,
                    "note": "BOM 정보가 없습니다."
                })
                continue

            min_cap_curr = float("inf")
            min_cap_avail = float("inf")
            bottleneck_part = ""

            for b in bom_items:
                part_inv = inv_map.get(b.part_name, {})
                curr_stock = part_inv.get("current_stock", 0.0)
                avail_stock = part_inv.get("available_stock", 0.0)

                if b.unit_qty > 0:
                    cap_curr = math.floor(curr_stock / b.unit_qty)
                    cap_avail = math.floor(avail_stock / b.unit_qty)
                else:
                    cap_curr = float("inf")
                    cap_avail = float("inf")

                if cap_curr < min_cap_curr:
                    min_cap_curr = cap_curr
                    bottleneck_part = b.part_name

                if cap_avail < min_cap_avail:
                    min_cap_avail = cap_avail

            max_cap = int(max(min_cap_curr, 0))
            potential_cap = int(max(min_cap_avail, 0))
            is_possible = (max_cap > 0)
            status_text = "정상 생산 가능" if is_possible else "생산 불가 (결품)"
            note = f"현재고 기준 {max_cap:,} Set 가능 (병목: {bottleneck_part})" if is_possible else f"생산 불가 ({bottleneck_part} 재고 부족)"

            capacity_rows.append({
                "fg_code": fg.code,
                "fg_name": fg.name,
                "spec": fg.spec,
                "max_capacity": max_cap,
                "bottleneck_part": bottleneck_part,
                "status": status_text,
                "potential_capacity": potential_cap,
                "note": note
            })

        return capacity_rows

    def simulate_plans(self, plans: Optional[Dict[str, float]] = None) -> Dict[str, Any]:
        """
        생산계획 시뮬레이션 및 부품별 계획 소요량/과부족 분석
        """
        if plans is not None:
            self.production_plans = plans

        inv_rows = self.calculate_inventory()
        cap_rows = {row["fg_code"]: row for row in self.calculate_capacity(inv_rows)}

        # 1. 완제품 계획 시뮬레이션 테이블
        fg_sim_list = []
        total_plan = 0.0
        total_max_cap = 0.0
        total_shortage = 0.0

        for fg in self.finished_goods.values():
            plan_qty = float(self.production_plans.get(fg.code, 0.0))
            cap_info = cap_rows.get(fg.code, {})
            max_cap = float(cap_info.get("max_capacity", 0))
            is_ok = (plan_qty <= max_cap)
            shortage = max(plan_qty - max_cap, 0.0) if plan_qty > 0 else 0.0

            total_plan += plan_qty
            total_max_cap += max_cap
            total_shortage += shortage

            if plan_qty == 0:
                sim_status = "-"
                sim_note = "계획 없음"
            elif is_ok:
                sim_status = "생산 가능"
                sim_note = "가용 재고 충족"
            else:
                sim_status = "생산 불가 (부족)"
                sim_note = f"가용량 대비 {int(shortage):,} Set 부족"

            fg_sim_list.append({
                "fg_code": fg.code,
                "fg_name": fg.name,
                "plan_qty": int(plan_qty),
                "max_capacity": int(max_cap),
                "status": sim_status,
                "shortage": int(shortage),
                "note": sim_note
            })

        # 2. 부품별 계획 총소요량 및 과부족 분석
        plan_req_by_part: Dict[str, float] = {}
        for b in self.bom_list:
            fg_plan = float(self.production_plans.get(b.fg_code, 0.0))
            plan_req_by_part[b.part_name] = plan_req_by_part.get(b.part_name, 0.0) + (b.unit_qty * fg_plan)

        part_sim_list = []
        total_plan_req = 0.0
        total_plan_shortage = 0.0

        for item in inv_rows:
            p_name = item["name"]
            curr_stock = item["current_stock"]
            plan_req = plan_req_by_part.get(p_name, 0.0)
            exp_rem = curr_stock - plan_req
            status = "부족" if exp_rem < 0 else "정상"
            add_order = abs(exp_rem) if exp_rem < 0 else 0.0

            total_plan_req += plan_req
            total_plan_shortage += add_order

            part_sim_list.append({
                "code": item["code"],
                "name": p_name,
                "category": item["category"],
                "unit": item["unit"],
                "current_stock": curr_stock,
                "plan_req": plan_req,
                "expected_remaining": exp_rem,
                "status": status,
                "additional_order_needed": add_order
            })

        return {
            "fg_plans": fg_sim_list,
            "part_requirements": part_sim_list,
            "summary": {
                "total_plan_qty": int(total_plan),
                "total_max_cap": int(total_max_cap),
                "total_fg_shortage": int(total_shortage),
                "total_part_req": int(total_plan_req),
                "total_part_shortage": int(total_plan_shortage)
            }
        }

    def get_kpi_summary(self) -> Dict[str, Any]:
        inv_rows = self.calculate_inventory()
        cap_rows = self.calculate_capacity(inv_rows)
        fg_inv_rows = self.calculate_fg_inventory()

        total_parts = len(inv_rows)
        shortage_parts = sum(1 for r in inv_rows if r["status"] == "부족")
        warning_parts = sum(1 for r in inv_rows if r["status"] == "주의")
        ready_fgs = sum(1 for r in cap_rows if r["max_capacity"] > 0)
        total_fgs = len(cap_rows)
        total_order_needed = sum(r["order_needed"] for r in inv_rows)
        pending_po_total = sum(r["pending_po"] for r in inv_rows)
        total_max_cap_sum = sum(r["max_capacity"] for r in cap_rows)
        fg_in_stock = sum(1 for r in fg_inv_rows if r["current_stock"] > 0)
        fg_out_stock = sum(1 for r in fg_inv_rows if r["current_stock"] <= 0)
        total_fg_stock = sum(r["current_stock"] for r in fg_inv_rows)

        return {
            "total_parts": total_parts,
            "shortage_parts": shortage_parts,
            "warning_parts": warning_parts,
            "ready_fgs": ready_fgs,
            "total_fgs": total_fgs,
            "total_order_needed": int(total_order_needed),
            "pending_po_total": int(pending_po_total),
            "total_max_cap_sum": int(total_max_cap_sum),
            "fg_in_stock": fg_in_stock,
            "fg_out_stock": fg_out_stock,
            "total_fg_stock": int(total_fg_stock)
        }
