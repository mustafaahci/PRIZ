import ctypes
import json
import os
import sys
import typing
from ctypes import wintypes

import PySide2
import win32api
import win32gui
from PySide2.QtCore import QRect, Qt
from PySide2.QtGui import QPixmap
from PySide2.QtWidgets import QLayout, QFrame, QVBoxLayout, QWidget, QApplication, QLineEdit, QLabel, QHBoxLayout
from PySide2.QtWinExtras import QtWin
from win32con import GWL_STYLE, WM_NCCALCSIZE, WS_CAPTION, HWND_TOPMOST, SWP_NOMOVE, SWP_NOSIZE, SWP_SHOWWINDOW


def readCss(path) -> str:
    with open(path, "r") as css:
        style = css.read()
    return style


def clearLayout(layout: QLayout) -> None:
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(0)


class ListWidget(QWidget):

    def __init__(self, name: str):
        super().__init__()
        self.name = name
        self.setStyleSheet("background: transparent;")
        # ELEMENTS
        self.app_name = QLabel()
        self.app_description = QLabel("description")
        self.app_icon = QLabel()

        # LAYOUTS
        self.vLayout = QVBoxLayout()
        self.hLayout = QHBoxLayout()

        # STAFF
        self.app_name.setStyleSheet("color: #FFFFFF; font: 14px Comfortaa")
        self.app_description.setStyleSheet("color: rgba(255,255,255,.5);")

        # ADD ELEMENTS TO LAYOUTS
        self.hLayout.addWidget(self.app_icon, 0)
        self.hLayout.addLayout(self.vLayout, 1)
        self.vLayout.addWidget(self.app_name)
        self.vLayout.addWidget(self.app_description)
        self.vLayout.addStretch()

        self.vLayout.setContentsMargins(15, 0, 0, 0)
        self.hLayout.setContentsMargins(16, 16, 16, 16)

        self.setLayout(self.hLayout)

    def setText(self, text, description):
        self.app_name.setText(text)
        self.app_description.setText(description)

    def setIcon(self, imagePath):
        self.app_icon.setPixmap(QPixmap(imagePath))


class MainFrame(QFrame):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("MainFrame")

        _layout = QVBoxLayout()
        _layout.setSpacing(0)
        _layout.setContentsMargins(20, 10, 20, 10)

        self.entry = QLineEdit()
        self.entry.setPlaceholderText("Please type 'exit' to exit!")
        self.entry.setObjectName("entry")

        _layout.addWidget(self.entry)

        self.setStyleSheet(readCss("style/main.css"))

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
        win32gui.SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0,
                              SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)

        if QtWin.isCompositionEnabled():
            # Don't working PySide2 < v: 5.15.1)
            QtWin.extendFrameIntoClientArea(self, -1, -1, -1, -1)
        else:
            QtWin.resetExtendedFrame(self)

        _layout = QVBoxLayout()
        clearLayout(_layout)

        _rect: QRect = QApplication.instance().desktop().availableGeometry(self)

        self._main_frame = MainFrame(self)
        _layout.addWidget(self._main_frame)

        # self.frame = QFrame(self)
        # self.frame.setStyleSheet("background: green;")
        # self.frame.setGeometry(0, 95, 250, 5)

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


def getPrograms():
    programList = []
    out = os.path.join(os.environ["ALLUSERSPROFILE"], "Start Menu", "Programs")

    for root, dirs, files in os.walk(out):
        for file in files:
            if file.endswith(".lnk"):
                programList.append(str(os.path.join(root, file)))

    with open('database/package.json', 'r', encoding='utf-8') as f:
        JSON = json.load(f)
        JSON["APPS"] = programList

    with open('database/package.json', 'w', encoding='utf-8') as f:
        json.dump(JSON, f, ensure_ascii=False, indent=4)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    WIDTH = 1000
    HEIGHT = 100

    # TODO: checkout external files
    # TODO: load database / package.json & settings.json

    try:
        with open('database/package.json', 'r', encoding='utf-8') as f:
            __JSON = json.load(f)
            if len(__JSON["APPS"]) == 0:
                getPrograms()
    except FileNotFoundError:
        with open('database/package.json', 'w'):
            pass
        getPrograms()

    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec_())
