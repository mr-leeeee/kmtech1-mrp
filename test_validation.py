import pytest
from mrp_engine import (MRPSystem, FinishedGood, validate_positive_qty,
                        validate_non_negative_qty, validate_fg_code_exists)


@pytest.fixture
def mrp():
    sys_ = MRPSystem()
    sys_.add_finished_good(FinishedGood(code="FG001", name="테스트완제품", spec="A",
                                        unit="Set", note=""))
    return sys_


class TestValidatePositiveQty:
    def test_positive_qty_ok(self):
        assert validate_positive_qty(10.5) == 10.5

    def test_zero_rejected(self):
        with pytest.raises(ValueError):
            validate_positive_qty(0)

    def test_negative_rejected(self):
        with pytest.raises(ValueError):
            validate_positive_qty(-3)

    def test_custom_field_name_in_message(self):
        with pytest.raises(ValueError, match="출하수량"):
            validate_positive_qty(0, "출하수량")


class TestValidateNonNegativeQty:
    def test_zero_allowed(self):
        assert validate_non_negative_qty(0) == 0

    def test_positive_allowed(self):
        assert validate_non_negative_qty(7) == 7

    def test_negative_rejected(self):
        with pytest.raises(ValueError):
            validate_non_negative_qty(-1)


class TestValidateFgCodeExists:
    def test_existing_fg_ok(self, mrp):
        assert validate_fg_code_exists(mrp, "FG001") == "FG001"

    def test_missing_fg_rejected(self, mrp):
        with pytest.raises(KeyError):
            validate_fg_code_exists(mrp, "FG999")


class TestAddFgTransactionValidation:
    def test_happy_path(self, mrp):
        tx = mrp.add_fg_transaction("2026-09-20", "SLIP-1", "FG001", 5.0,
                                    customer="거래처A", worker="김")
        assert tx.out_qty == 5.0
        assert tx.fg_name == "테스트완제품"

    def test_zero_out_qty_rejected(self, mrp):
        with pytest.raises(ValueError):
            mrp.add_fg_transaction("2026-09-20", "SLIP-2", "FG001", 0,
                                   customer="거래처A")

    def test_unknown_fg_rejected(self, mrp):
        with pytest.raises(KeyError):
            mrp.add_fg_transaction("2026-09-20", "SLIP-3", "FG999", 1,
                                   customer="거래처A")


class TestUpdateFgTransactionValidation:
    def test_update_ok(self, mrp):
        tx = mrp.add_fg_transaction("2026-09-20", "SLIP-1", "FG001", 5.0,
                                    customer="A")
        updated = mrp.update_fg_transaction(tx.id, "2026-09-21", "SLIP-1", "FG001",
                                            7.0, customer="A")
        assert updated.out_qty == 7.0

    def test_update_negative_rejected(self, mrp):
        tx = mrp.add_fg_transaction("2026-09-20", "SLIP-1", "FG001", 5.0,
                                    customer="A")
        with pytest.raises(ValueError):
            mrp.update_fg_transaction(tx.id, "2026-09-21", "SLIP-1", "FG001",
                                      -1.0, customer="A")

    def test_update_missing_id_rejected(self, mrp):
        with pytest.raises(KeyError):
            mrp.update_fg_transaction(99999, "2026-09-21", "SLIP-1", "FG001",
                                      7.0, customer="A")


class TestAddProductionLogValidation:
    def test_happy_path(self, mrp):
        log = mrp.add_production_log("2026-09-20", "P-1", "L1", "FG001",
                                     prod_qty=10, good_qty=9, defect_qty=1,
                                     worker="박")
        assert log.prod_qty == 10

    def test_qty_mismatch_rejected(self, mrp):
        with pytest.raises(ValueError, match="수량 불일치"):
            mrp.add_production_log("2026-09-20", "P-2", "L1", "FG001",
                                   prod_qty=10, good_qty=8, defect_qty=1,
                                   worker="박")

    def test_negative_qty_rejected(self, mrp):
        with pytest.raises(ValueError, match="음수"):
            mrp.add_production_log("2026-09-20", "P-3", "L1", "FG001",
                                   prod_qty=-5, good_qty=-5, defect_qty=0,
                                   worker="박")