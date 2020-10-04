import os
import sys
import json
import typing
import ctypes
from ctypes import wintypes
from time import time, sleep

from PySide2.QtCore import QRect, Qt, QSize, QFileInfo
from PySide2.QtGui import QPixmap, QPalette
from PySide2.QtWidgets import QLayout, QFrame, QVBoxLayout, QWidget, QApplication, QLineEdit, QLabel, QHBoxLayout, \
    QListWidget, QListWidgetItem, QFileIconProvider, QFileSystemModel, QSplashScreen

import win32api
import win32gui
import win32com.client
from PySide2.QtWinExtras import QtWin
from win32con import GWL_STYLE, WM_NCCALCSIZE, WS_CAPTION, HWND_TOPMOST, SWP_NOMOVE, SWP_NOSIZE, SWP_SHOWWINDOW


def readCss(path) -> str:
    with open(path, "r") as css:
        style = css.read()
    return style


def clearLayout(layout: QLayout) -> QLayout:
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(0)
    return layout


class SplashScreen(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setFixedSize(1040, 140)
        self.setAttribute(Qt.WA_DeleteOnClose)
        win32gui.SetWindowPos(self.winId(), HWND_TOPMOST, 0, 0, 0, 0,
                              SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)

        __background = QPixmap("image/SplashScreen.png")
        __background.scaled(self.size())
        __palette = QPalette()
        __palette.setBrush(QPalette.Background, __background)
        self.setPalette(__palette)


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
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.provider = QFileIconProvider()

        _layout = QVBoxLayout()
        _layout.setSpacing(20)
        _layout.setContentsMargins(20, 20, 20, 20)

        self.entry = QLineEdit()
        self.entry.setPlaceholderText("Please type 'exit' to exit!")
        self.entry.setObjectName("entry")

        self.result_list = QListWidget()
        self.result_list.setObjectName("result_list")
        self.result_list.itemDoubleClicked.connect(self.openProgram)

        _layout.addWidget(self.entry)
        _layout.addWidget(self.result_list)
        self.setStyleSheet(readCss("style/main.css"))
        self.setLayout(_layout)
        self.getApps()

    def openProgram(self):
        os.startfile(self.result_list.currentItem().text())

    @staticmethod
    def getAppIcon(path):
        fin = QFileInfo(path)
        model = QFileSystemModel()

        model.setRootPath(fin.path())
        qq = model.iconProvider()
        icon = qq.icon(fin)
        return icon

    def getApps(self):
        for app_path in PACKAGE["APPS"]:
            app_name = str(os.path.basename(app_path)).split(".")[0]
            print(app_path)

            new_widget = ListWidget(app_name)
            new_widget.setText(app_name, app_path)

            icon = self.getAppIcon(app_path)

            pixmap = icon.pixmap(icon.actualSize(QSize(32, 32)))
            new_widget.setIcon(pixmap)

            new_item = QListWidgetItem(app_path, self.result_list)
            new_item.setSizeHint(new_widget.sizeHint())
            self.result_list.addItem(new_item)
            self.result_list.setItemWidget(new_item, new_widget)


class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setObjectName("MainWindow")
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowMinimizeButtonHint)

        self.__press_pos = None

        _layout = clearLayout(QVBoxLayout())
        self._main_frame = MainFrame(self)
        _layout.addWidget(self._main_frame)

        # self.frame = QFrame(self)
        # self.frame.setStyleSheet("background: green;")
        # self.frame.setGeometry(0, 95, 250, 5)

        self.show()

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

        _rect: QRect = QApplication.instance().desktop().availableGeometry(self)

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

    def nativeEvent(self, eventType, message: int) -> typing.Tuple:
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

    with open('database/package.json', 'r', encoding='utf-8') as file:
        JSON = json.load(file)
        JSON["APPS"] = programList

    with open('database/package.json', 'w', encoding='utf-8') as file:
        json.dump(JSON, file, ensure_ascii=False, indent=4)

    return JSON


if __name__ == "__main__":
    app = QApplication(sys.argv)

    WIDTH = 1000
    HEIGHT = 100
    PACKAGE = None

    # TODO: checkout external files
    # TODO: load database / package.json & settings.json

    try:
        with open('database/package.json', 'r', encoding='utf-8') as f:
            __JSON = json.load(f)
            if len(__JSON["APPS"]) == 0:
                PACKAGE = getPrograms()
            else:
                PACKAGE = __JSON
    except FileNotFoundError:
        with open('database/package.json', 'w'):
            pass
        PACKAGE = getPrograms()

    # splash_screen = SplashScreen()
    start = time()
    splash = QSplashScreen(QPixmap("image/SplashScreen.png"), Qt.WindowStaysOnTopHint)
    splash.show()
    while time() - start < 1:
        sleep(0.001)
        app.processEvents()
    MainWindow = MainWindow()
    splash.finish(MainWindow)

    sys.exit(app.exec_())
