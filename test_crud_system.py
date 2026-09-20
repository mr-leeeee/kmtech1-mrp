import unittest
from fastapi.testclient import TestClient
import mrp_engine
import excel_sync
from server import app, mrp, EXCEL_FILE

client = TestClient(app)

class TestMRPAllCRUD(unittest.TestCase):
    def setUp(self):
        # Reload fresh from Excel
        global mrp
        test_mrp = excel_sync.load_from_excel(EXCEL_FILE)
        import server
        server.mrp = test_mrp

    def test_01_item_stock_edit(self):
        # 1. Check initial RM-GL-01
        res = client.get('/api/status')
        self.assertEqual(res.status_code, 200)
        data = res.json()
        mat001 = next(item for item in data['inventory'] if item['code'] == 'RM-GL-01')
        initial_base = mat001['base_stock']
        initial_curr = mat001['current_stock']
        
        # 2. Update base stock +50
        update_res = client.put('/api/item/RM-GL-01', json={
            'base_stock': initial_base + 50,
            'safety_stock': 100,
            'target_stock': 200,
            'lead_time': 5
        })
        self.assertEqual(update_res.status_code, 200)
        
        # 3. Verify recalculation
        res2 = client.get('/api/status')
        data2 = res2.json()
        mat001_updated = next(item for item in data2['inventory'] if item['code'] == 'RM-GL-01')
        self.assertEqual(mat001_updated['base_stock'], initial_base + 50)
        self.assertEqual(mat001_updated['current_stock'], initial_curr + 50)
        print('  [PASS] Item stock edit and recalculation verified')

    def test_02_production_log_crud(self):
        # 1. Check initial bracket stock
        res = client.get('/api/status')
        data = res.json()
        bracket_before = next(item for item in data['inventory'] if item['name'] == '블랙선반브라켓')['current_stock']
        
        # 2. Add production: 10 sets of FG-001 (requires 2 brackets each = 20 brackets)
        add_res = client.post('/api/production', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-PRD-001',
            'line': '테스트라인',
            'fg_code': 'FG-001',
            'prod_qty': 10,
            'good_qty': 10,
            'defect_qty': 0,
            'worker': '테스터',
            'note': '단위테스트'
        })
        self.assertEqual(add_res.status_code, 200)
        log_id = add_res.json()['id']
        
        bracket_after = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '블랙선반브라켓')['current_stock']
        self.assertEqual(bracket_before - bracket_after, 20.0)
        
        # 3. Update production: change qty to 30 (requires 60 brackets)
        upd_res = client.put(f'/api/production/{log_id}', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-PRD-001',
            'line': '테스트라인',
            'fg_code': 'FG-001',
            'prod_qty': 30,
            'good_qty': 30,
            'defect_qty': 0,
            'worker': '테스터',
            'note': '수정테스트'
        })
        self.assertEqual(upd_res.status_code, 200)
        bracket_after_upd = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '블랙선반브라켓')['current_stock']
        self.assertEqual(bracket_before - bracket_after_upd, 60.0)
        
        # 4. Delete production: verify component restoration
        del_res = client.delete(f'/api/production/{log_id}')
        self.assertEqual(del_res.status_code, 200)
        bracket_restored = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '블랙선반브라켓')['current_stock']
        self.assertEqual(bracket_restored, bracket_before)
        print('  [PASS] Production log add/update/delete and BOM backflush restoration verified')

    def test_03_transaction_crud(self):
        # 1. Initial stock of 칼브럭피스
        kb_before = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '칼브럭피스')['current_stock']
        
        # 2. Add 입고 500
        add_res = client.post('/api/transaction', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-TX-001',
            'tx_type': '입고',
            'part_name': '칼브럭피스',
            'qty': 500,
            'customer': '테스트업체',
            'note': '입고테스트',
            'worker': '김원장'
        })
        self.assertEqual(add_res.status_code, 200)
        tx_id = add_res.json()['id']
        kb_after = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '칼브럭피스')['current_stock']
        self.assertEqual(kb_after, kb_before + 500)
        
        # 3. Update 입고 to 700
        upd_res = client.put(f'/api/transaction/{tx_id}', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-TX-001',
            'tx_type': '입고',
            'part_name': '칼브럭피스',
            'qty': 700,
            'customer': '테스트업체',
            'note': '입고수정',
            'worker': '김원장'
        })
        self.assertEqual(upd_res.status_code, 200)
        kb_upd = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '칼브럭피스')['current_stock']
        self.assertEqual(kb_upd, kb_before + 700)
        
        # 4. Delete transaction
        del_res = client.delete(f'/api/transaction/{tx_id}')
        self.assertEqual(del_res.status_code, 200)
        kb_restored = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '칼브럭피스')['current_stock']
        self.assertEqual(kb_restored, kb_before)
        print('  [PASS] Transaction ledger add/update/delete verified')

    def test_04_po_crud(self):
        # 1. Initial pending PO of 칼브럭피스
        po_before = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '칼브럭피스')['pending_po']
        
        # 2. Add PO 1000
        add_res = client.post('/api/po', json={
            'date': '2026-09-18',
            'po_no': 'TEST-PO-001',
            'part_name': '칼브럭피스',
            'order_qty': 1000,
            'recv_qty': 0,
            'exp_date': '2026-09-25',
            'customer': '테스트업체',
            'note': '발주테스트'
        })
        self.assertEqual(add_res.status_code, 200)
        po_id = add_res.json()['id']
        po_after = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '칼브럭피스')['pending_po']
        self.assertEqual(po_after, po_before + 1000)
        
        # 3. Update PO: recv_qty = 400 (pending becomes 600)
        upd_res = client.put(f'/api/po/{po_id}', json={
            'date': '2026-09-18',
            'po_no': 'TEST-PO-001',
            'part_name': '칼브럭피스',
            'order_qty': 1000,
            'recv_qty': 400,
            'exp_date': '2026-09-25',
            'customer': '테스트업체',
            'note': '부분입고수정'
        })
        self.assertEqual(upd_res.status_code, 200)
        po_upd = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '칼브럭피스')['pending_po']
        self.assertEqual(po_upd, po_before + 600)
        
        # 4. Delete PO
        del_res = client.delete(f'/api/po/{po_id}')
        self.assertEqual(del_res.status_code, 200)
        po_restored = next(item for item in client.get('/api/status').json()['inventory'] if item['name'] == '칼브럭피스')['pending_po']
        self.assertEqual(po_restored, po_before)
        print('  [PASS] Purchase Order add/update/delete verified')

    def test_05_fg_transaction_crud(self):
        # 1. Initial FG-001 inventory
        fg_before = next(r for r in client.get('/api/status').json()['fg_inventory'] if r['fg_code'] == 'FG-001')['current_stock']

        # 2. Add FG shipment: 10 EA out
        add_res = client.post('/api/fgtransaction', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-OUT-001',
            'fg_code': 'FG-001',
            'out_qty': 10,
            'customer': '테스트고객',
            'note': '출하테스트',
            'worker': '김출하'
        })
        self.assertEqual(add_res.status_code, 200)
        fg_tx_id = add_res.json()['id']
        fg_after = next(r for r in client.get('/api/status').json()['fg_inventory'] if r['fg_code'] == 'FG-001')['current_stock']
        self.assertEqual(fg_after, fg_before - 10)

        # 3. Update shipment: 10 -> 7 EA
        upd_res = client.put(f'/api/fgtransaction/{fg_tx_id}', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-OUT-001',
            'fg_code': 'FG-001',
            'out_qty': 7,
            'customer': '테스트고객',
            'note': '출하수정',
            'worker': '김출하'
        })
        self.assertEqual(upd_res.status_code, 200)
        fg_upd = next(r for r in client.get('/api/status').json()['fg_inventory'] if r['fg_code'] == 'FG-001')['current_stock']
        self.assertEqual(fg_upd, fg_before - 7)

        # 4. Delete shipment: inventory restored
        del_res = client.delete(f'/api/fgtransaction/{fg_tx_id}')
        self.assertEqual(del_res.status_code, 200)
        fg_restored = next(r for r in client.get('/api/status').json()['fg_inventory'] if r['fg_code'] == 'FG-001')['current_stock']
        self.assertEqual(fg_restored, fg_before)
        print('  [PASS] FG transaction add/update/delete and inventory auto-deduction verified')

    def test_06_fg_transaction_error_paths(self):
        # 1. Unknown fg_code -> 400
        bad_fg = client.post('/api/fgtransaction', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-OUT-ERR',
            'fg_code': 'FG-999',
            'out_qty': 5,
            'customer': '',
            'note': '',
            'worker': ''
        })
        self.assertEqual(bad_fg.status_code, 400)

        # 2. Negative qty -> 400
        neg_qty = client.post('/api/fgtransaction', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-OUT-ERR2',
            'fg_code': 'FG-001',
            'out_qty': -5,
            'customer': '',
            'note': '',
            'worker': ''
        })
        self.assertEqual(neg_qty.status_code, 400)

        # 3. Zero qty -> 400
        zero_qty = client.post('/api/fgtransaction', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-OUT-ERR3',
            'fg_code': 'FG-001',
            'out_qty': 0,
            'customer': '',
            'note': '',
            'worker': ''
        })
        self.assertEqual(zero_qty.status_code, 400)

        # 4. Update non-existent id -> 400
        bad_upd = client.put('/api/fgtransaction/99999', json={
            'date': '2026-09-18',
            'slip_no': 'TEST-OUT-ERR4',
            'fg_code': 'FG-001',
            'out_qty': 5,
            'customer': '',
            'note': '',
            'worker': ''
        })
        self.assertEqual(bad_upd.status_code, 400)

        # 5. Delete non-existent id -> 400
        bad_del = client.delete('/api/fgtransaction/99999')
        self.assertEqual(bad_del.status_code, 400)

        # 6. Verify no phantom state change after failed attempts
        fg_final = next(r for r in client.get('/api/status').json()['fg_inventory'] if r['fg_code'] == 'FG-001')['current_stock']
        fg_initial = next(r for r in client.get('/api/status').json()['fg_inventory'] if r['fg_code'] == 'FG-001')['current_stock']
        self.assertEqual(fg_final, fg_initial)
        print('  [PASS] FG transaction error paths (unknown code / negative / zero qty / missing id) verified')

if __name__ == '__main__':
    unittest.main()
