from unittest.mock import Mock, call, patch

import pytest

from report_automation_engine.hwp_com_writer import (
    HwpWriterError,
    connect_hwp,
    register_file_path_checker,
    write_hwp_document,
)


RESULTS = (
    "FilePathCheckerModuleExample",
    "FilePathCheckerModule",
    "FilePathCheckDLL",
)


def test_hancom_example_module_is_tried_first():
    hwp = Mock()
    hwp.RegisterModule.return_value = True
    report = {"com": {}, "warnings": []}

    register_file_path_checker(hwp, report)

    hwp.RegisterModule.assert_called_once_with("FilePathCheckDLL", RESULTS[0])
    assert report["com"]["file_path_checker"] == RESULTS[0]
    assert not report["warnings"]


def test_false_and_exception_do_not_mask_working_legacy_module():
    hwp = Mock()
    hwp.RegisterModule.side_effect = [False, RuntimeError("not installed"), True]
    report = {"com": {}, "warnings": []}

    register_file_path_checker(hwp, report)

    assert hwp.RegisterModule.call_args_list == [call("FilePathCheckDLL", name) for name in RESULTS]
    assert report["com"]["file_path_checker"] == RESULTS[2]
    assert [step["module"] for step in report["com"]["steps"] if step["status"] == "failed"] == list(RESULTS[:2])


def test_all_false_results_are_reported_as_unregistered():
    hwp = Mock()
    hwp.RegisterModule.return_value = False
    report = {"com": {"file_path_checker": ""}, "warnings": []}

    register_file_path_checker(hwp, report)

    assert not report["com"]["file_path_checker"]
    assert report["warnings"]
    assert not any(step["status"] == "ready" for step in report["com"]["steps"])


@pytest.mark.parametrize("native_ready, pyhwpx_ready, success", [
    (True, False, True),
    (False, True, True),
    (False, False, False),
])
def test_pyhwpx_uses_existing_registration_or_explicitly_sets_it_up(native_ready, pyhwpx_ready, success):
    session = Mock()
    session.Application.RegisterModule.return_value = native_ready
    session.register_module.return_value = pyhwpx_ready
    package = Mock(Hwp=Mock(return_value=session))
    report = {"com": {}, "warnings": []}
    with patch.dict("sys.modules", {"pyhwpx": package}):
        hwp, owner = connect_hwp(report, None, "pyhwpx")

    package.Hwp.assert_called_once_with(new=True, visible=False, register_module=False)
    assert hwp is session.Application
    assert owner is session
    assert session.register_module.call_count == (0 if native_ready else 1)
    assert bool(report["com"]["file_path_checker"]) is success
    assert bool(report["warnings"]) is not success


def test_missing_pyhwpx_has_actionable_error():
    with patch.dict("sys.modules", {"pyhwpx": None}):
        with pytest.raises(HwpWriterError, match="pip install pyhwpx"):
            connect_hwp({"com": {}, "warnings": []}, None, "pyhwpx")


def test_hidden_writer_stops_before_open_if_module_is_missing(tmp_path):
    package, preflight, template, output = [
        tmp_path / name for name in ("package.json", "preflight.json", "template.hwpx", "output.hwpx")
    ]
    package.write_text("{}", encoding="utf-8")
    preflight.write_text('{"status":"ready"}', encoding="utf-8")
    template.touch()
    hwp = Mock()
    with (
        patch("report_automation_engine.hwp_com_writer.platform.system", return_value="Windows"),
        patch("report_automation_engine.hwp_com_writer.acquire_writer_lock", return_value=None),
        patch("report_automation_engine.hwp_com_writer.register_styles_or_raise"),
        patch("report_automation_engine.hwp_com_writer.connect_hwp", return_value=(hwp, None)),
        pytest.raises(HwpWriterError) as error,
    ):
        write_hwp_document(package, preflight, template, output)
    assert error.value.action == "file_path_checker"
    hwp.Open.assert_not_called()
    hwp.Quit.assert_called_once()
    assert not output.exists()
