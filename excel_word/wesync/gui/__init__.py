"""桌面界面（PySide6）：把各功能内置成页面，左侧导航切换。

启动方式：
    python -m wesync.gui          或双击  审计工具箱.pyw
"""

from __future__ import annotations

from pathlib import Path


def run(base_dir: Path | None = None) -> int:
    """启动界面主循环，返回进程退出码。"""
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QIcon
    from PySide6.QtWidgets import QApplication

    from . import theme
    from .app import WesyncApp
    from .resources import icon_path

    # 按系统实际缩放比例渲染（150% 即 1.5 倍）。不能用 Round：它会把 150% 取整成 2 倍，
    # 字体在一部分高分屏电脑上显得异常大；Qt6 支持分数缩放，不必担心半像素发虚。
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication.instance() or QApplication([])
    app.setApplicationName("审计工具箱")
    app.setWindowIcon(QIcon(str(icon_path())))
    theme.apply(app)
    window = WesyncApp(base_dir=base_dir)
    window.show()
    return app.exec()


__all__ = ["run"]