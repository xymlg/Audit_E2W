import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest  # noqa: E402

# GUI 测试隔离：偏好文件（.wesync_ui.json）默认写项目根，会被真实使用记录污染，
# 也让测试反过来覆盖用户偏好。改为指向测试工作目录内的独立文件。
from helpers import WORK  # noqa: E402
from wesync import report_format as _report_format  # noqa: E402
from wesync.gui import store as _ui_store  # noqa: E402

_ui_store._config_path = lambda: WORK / ".wesync_ui_test.json"


@pytest.fixture(autouse=True)
def _isolate_format_profiles(monkeypatch, request):
    profiles = WORK / "format_profiles" / request.node.name
    profiles.mkdir(parents=True, exist_ok=True)
    monkeypatch.setattr(_report_format, "_profiles_dir", lambda: profiles)


@pytest.fixture(scope="session")
def qapp():
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture(autouse=True)
def _restore_log_sink():
    # GUI 任务结束会把全局日志 sink 置 None；每个测试前恢复到 LOGS，保证日志断言隔离。
    from helpers import LOGS
    from wesync import console
    console.set_log_sink(LOGS.append)
    yield