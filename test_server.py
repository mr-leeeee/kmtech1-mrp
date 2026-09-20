import sys
from fastapi.testclient import TestClient
from server import app

client = TestClient(app)

print("=== 1. TEST GET /api/status ===")
res = client.get("/api/status")
assert res.status_code == 200, f"Status code: {res.status_code}"
data = res.json()
kpis = data["kpis"]
print("KPIs:", kpis)
assert kpis["total_parts"] == 14
assert kpis["ready_fgs"] == 6
assert kpis["total_fgs"] == 6
print("GET /api/status PASSED!")

print("\n=== 2. TEST POST /api/production (FG-001 10 EA) ===")
# Base stock for 블랙삼각선반
inv_before = {item["name"]: item for item in data["inventory"]}
tri_curr_before = inv_before["블랙삼각선반"]["current_stock"]

res_prod = client.post("/api/production", json={
    "date": "2026-09-19",
    "slip_no": "PRD-TEST-01",
    "line": "조립 1라인",
    "fg_code": "FG-001",
    "prod_qty": 10,
    "good_qty": 10,
    "defect_qty": 0,
    "worker": "테스터",
    "note": "자동화 테스트"
})
assert res_prod.status_code == 200, res_prod.text
print("Prod Response:", res_prod.json())

# Verify backflush
res2 = client.get("/api/status")
inv_after = {item["name"]: item for item in res2.json()["inventory"]}
tri_curr_after = inv_after["블랙삼각선반"]["current_stock"]
print(f"블랙삼각선반 현재고: {tri_curr_before} -> {tri_curr_after} (차감 10 EA)")
assert tri_curr_after == tri_curr_before - 10, f"Expected {tri_curr_before - 10}, got {tri_curr_after}"
print("POST /api/production PASSED!")

print("\n=== 3. TEST POST /api/transaction (입고 100 EA) ===")
res_tx = client.post("/api/transaction", json={
    "date": "2026-09-19",
    "slip_no": "IN-TEST-01",
    "tx_type": "입고",
    "part_name": "블랙삼각선반",
    "qty": 100,
    "customer": "(주)대한유리",
    "worker": "김자재",
    "note": "보충 입고 테스트"
})
assert res_tx.status_code == 200, res_tx.text
res3 = client.get("/api/status")
inv_after_tx = {item["name"]: item for item in res3.json()["inventory"]}
tri_curr_after_tx = inv_after_tx["블랙삼각선반"]["current_stock"]
print(f"블랙삼각선반 현재고: {tri_curr_after} -> {tri_curr_after_tx} (입고 +100 EA)")
assert tri_curr_after_tx == tri_curr_after + 100
print("POST /api/transaction PASSED!")

print("\n=== 4. TEST POST /api/po (발주 50 EA) ===")
res_po = client.post("/api/po", json={
    "date": "2026-09-19",
    "po_no": "PO-TEST-01",
    "part_name": "블랙삼각선반",
    "order_qty": 50,
    "exp_date": "2026-09-25",
    "customer": "(주)대한유리",
    "note": "신규 발주 테스트"
})
assert res_po.status_code == 200, res_po.text
res4 = client.get("/api/status")
inv_after_po = {item["name"]: item for item in res4.json()["inventory"]}
tri_pend_after_po = inv_after_po["블랙삼각선반"]["pending_po"]
print(f"블랙삼각선반 미입고량: {tri_pend_after_po} EA (기존 30 + 신규 50 = 80 EA)")
assert tri_pend_after_po == 80
print("POST /api/po PASSED!")

print("\n=== 5. TEST POST /api/plan (생산계획 시뮬레이션) ===")
res_plan = client.post("/api/plan", json={
    "plans": {
        "FG-001": 200,
        "FG-002": 50,
        "FG-003": 0,
        "FG-004": 0,
        "FG-005": 0,
        "FG-006": 0
    }
})
assert res_plan.status_code == 200, res_plan.text
sim_data = res_plan.json()["simulation"]
print("Plan summary:", sim_data["summary"])
assert sim_data["summary"]["total_plan_qty"] == 250
print("POST /api/plan PASSED!")

print("\n=== 6. TEST Root / and /static/index.html ===")
res_root = client.get("/")
assert res_root.status_code == 200
assert "KMTECH" in res_root.text
print("GET / (Frontend HTML) PASSED!")

print("\nALL BACKEND & API TESTS COMPLETED SUCCESSFULLY!")
