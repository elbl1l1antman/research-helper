from unittest.mock import Mock, call

from report_automation_engine.hwp_com_writer import register_file_path_checker


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
