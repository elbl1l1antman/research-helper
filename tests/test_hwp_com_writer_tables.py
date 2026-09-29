from report_automation_engine.hwp_com_writer import (
    find_hwp_table_control,
    is_hwp_table_control,
    same_hwp_control,
    table_rows_for_hwp,
)


class FakeControl:
    def __init__(self, ctrl_id, instance_id=0):
        self.CtrlID = ctrl_id
        self.instance_id = instance_id

    def GetCtrlInstID(self):
        return self.instance_id


def test_table_rows_for_hwp_uses_cell_contract_display_text():
    table = {
        "matrix": [
            [
                {"display_text": "구분", "raw_value": "구분"},
                {"display_text": "비율", "raw_value": "비율"},
            ],
            [
                {"display_text": "전체", "raw_value": "전체"},
                {"display_text": "63.3%", "raw_value": 0.633335353},
            ],
        ]
    }

    assert table_rows_for_hwp(table) == [["구분", "비율"], ["전체", "63.3%"]]


def test_table_rows_for_hwp_does_not_truncate_wide_table():
    table = {
        "matrix": [
            [{"display_text": f"배너 {col}"} for col in range(1, 37)],
            [{"display_text": f"{col}.0"} for col in range(1, 37)],
        ]
    }

    rows = table_rows_for_hwp(table)

    assert len(rows) == 2
    assert len(rows[0]) == 36
    assert rows[0][-1] == "배너 36"
    assert rows[1][-1] == "36.0"


def test_hwp_table_control_requires_tbl_control_id():
    assert is_hwp_table_control(FakeControl("tbl"))
    assert not is_hwp_table_control(FakeControl("gso"))
    assert not is_hwp_table_control(None)


def test_find_hwp_table_control_checks_current_parent_and_last_controls():
    class FakeHwp:
        CurSelectedCtrl = FakeControl("gso")
        ParentCtrl = FakeControl("tbl")
        LastCtrl = FakeControl("gso")

    assert find_hwp_table_control(FakeHwp()).CtrlID == "tbl"


def test_same_hwp_control_uses_control_instance_id():
    assert same_hwp_control(FakeControl("tbl", 10), FakeControl("tbl", 10))
    assert not same_hwp_control(FakeControl("tbl", 10), FakeControl("tbl", 11))
