import sys
import json
import ctypes
from ctypes import wintypes


import typing
import PySide2
from PySide2.QtCore import QRect, Qt
from PySide2.QtWidgets import QLayout, QFrame, QVBoxLayout, QWidget, QApplication, QLineEdit
from PySide2.QtWinExtras import QtWin

import win32api
import win32gui
from win32con import GWL_STYLE, WM_NCCALCSIZE, WS_CAPTION


def clearLayout(layout: QLayout) -> None:
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(0)


class MainFrame(QFrame):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("MainFrame")

        _layout = QVBoxLayout()
        _layout.setSpacing(0)
        _layout.setContentsMargins(20, 10, 20, 10)

        self.entry = QLineEdit()
        self.entry.setPlaceholderText("Please type `exit` to exit!")
        self.entry.setObjectName("entry")

        _layout.addWidget(self.entry)

        with open("style/main.css", "r") as css:
            style = css.read()
            self.setStyleSheet(style)

        self.setLayout(_layout)


class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setObjectName("MainWindow")
        self.setWindowFlags(Qt.FramelessWindowHint
                            | Qt.WindowMinimizeButtonHint)

        self.__press_pos = None

        hWnd = self.winId()
        style = win32gui.GetWindowLong(hWnd, GWL_STYLE)
        win32gui.SetWindowLong(hWnd, GWL_STYLE, style | WS_CAPTION)

        if QtWin.isCompositionEnabled():
            # Don't working PySide2 < v: 5.15.1)
            QtWin.extendFrameIntoClientArea(self, -1, -1, -1, -1)
        else:
            QtWin.resetExtendedFrame(self)

        _layout = QVBoxLayout()
        clearLayout(_layout)

        _rect: QRect = app.instance().desktop().availableGeometry(self)

        self._main_frame = MainFrame(self)
        _layout.addWidget(self._main_frame)

        x = _rect.width() / 2 - WIDTH / 2
        y = _rect.height() / 2 - HEIGHT / 2 - _rect.height() * 15 / 100
        self.setGeometry(x, y, WIDTH, HEIGHT)

        self.setLayout(_layout)

    def getMainFrame(self) -> QFrame:
        return self._main_frame

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.__press_pos = event.pos()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.__press_pos = None

    def mouseMoveEvent(self, event):
        if self.__press_pos:
            self.move(self.pos() + (event.pos() - self.__press_pos))

    def nativeEvent(self, eventType: PySide2.QtCore.QByteArray, message: int) -> typing.Tuple:
        retval, result = super().nativeEvent(eventType, message)

        if eventType == "windows_generic_MSG":
            msg = wintypes.MSG.from_address(message.__int__())

            x = win32api.LOWORD(ctypes.c_long(msg.lParam).value) - self.frameGeometry().x()
            y = win32api.HIWORD(ctypes.c_long(msg.lParam).value) - self.frameGeometry().y()

            if self.childAt(x, y) is not None:
                return retval, result

            if msg.message == WM_NCCALCSIZE:
                return True, 0

        return retval, result


if __name__ == "__main__":
    app = QApplication(sys.argv)

    WIDTH = 1000
    HEIGHT = 100

    # TODO: checkout external files
    # TODO: load database / package.json & settings.json

    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec_())
