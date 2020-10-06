import ctypes
import json
import os
import sys
import typing
from ctypes import wintypes
from time import time, sleep

import win32api
import win32gui
from PySide2.QtCore import QRect, Qt, QSize, QFileInfo, QRegExp, QThread, Signal
from PySide2.QtGui import QPixmap, QPalette, QKeyEvent, QTextCharFormat, QSyntaxHighlighter, QColor, QFont
from PySide2.QtWidgets import QLayout, QFrame, QVBoxLayout, QWidget, QApplication, QLabel, QHBoxLayout, \
    QListWidget, QListWidgetItem, QFileSystemModel, QSplashScreen, QTextEdit, QAction, QMenu, \
    QSystemTrayIcon, QStyle
from PySide2.QtWinExtras import QtWin
from pynput.keyboard import Key, Listener
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
        self.app_name.setStyleSheet("color: #FFFFFF; font: 18px Rajdhani")
        self.app_description.setStyleSheet("color: rgba(255,255,255,.5); font: 16px Rajdhani")

        # ADD ELEMENTS TO LAYOUTS
        self.hLayout.addWidget(self.app_icon, 0)
        self.hLayout.addLayout(self.vLayout, 1)
        self.vLayout.addWidget(self.app_name)
        self.vLayout.addWidget(self.app_description)
        self.vLayout.addStretch()

        self.vLayout.setContentsMargins(15, 0, 0, 0)
        self.hLayout.setContentsMargins(16, 16, 16, 16)

        self.setLayout(self.hLayout)

    def getText(self):
        return self.app_name.text()

    def getDescription(self):
        return self.app_description.text()

    def setText(self, text, description):
        self.app_name.setText(text)
        if len(description) < 120:
            self.app_description.setText(description)
        else:
            self.app_description.setText(description[:120] + "...")

    def setIcon(self, imagePath):
        self.app_icon.setPixmap(QPixmap(imagePath))


class TextEdit(QTextEdit):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("entry")
        self.setPlaceholderText("Please type 'exit' to exit!")
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setLineWrapMode(QTextEdit.NoWrap)
        self.setFixedHeight(41)

    def keyPressEvent(self, e: QKeyEvent):
        if e.key() == Qt.Key_Return:
            self.parent().openProgram()
        else:
            super().keyPressEvent(e)


class Highlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super(Highlighter, self).__init__(parent)

        keywordFormat = QTextCharFormat()
        keywordFormat.setForeground(QColor("#00E5FF"))
        keywordFormat.setFontWeight(QFont.DemiBold)

        keywordPatterns = ["\\b" + syntax + "\\b" for syntax in SYNTAXS]

        self.highlightingRules = [(QRegExp(pattern), keywordFormat)
                                  for pattern in keywordPatterns]

        # singleLineCommentFormat = QTextCharFormat()
        # singleLineCommentFormat.setForeground(Qt.red)
        # self.highlightingRules.append((QRegExp("//[^\n]*"), singleLineCommentFormat))

    def highlightBlock(self, text):
        for pattern, _format in self.highlightingRules:
            expression = QRegExp(pattern)
            index = expression.indexIn(text)
            while index >= 0:
                length = expression.matchedLength()
                self.setFormat(index, length, _format)
                index = expression.indexIn(text, index + length)

        self.setCurrentBlockState(0)


# noinspection PyUnresolvedReferences
class KeyThread(QThread):
    """ KEY LISTENER """
    update = Signal(bool)
    tray_icon = Signal(bool)
    current_row = Signal(int)

    def __init__(self, parent=None):
        QThread.__init__(self)
        self.setParent(parent)
        self.result_list = self.parent().getResultList()
        self.entry = self.parent().getEntry()

    def run(self):
        """ WORKING """
        # The key combination to check
        COMBINATION = {Key.ctrl_l, Key.space}
        # The currently active modifiers
        current = set()

        def on_press(key):
            """ KEY LISTENER FUNC"""
            if key in COMBINATION:
                current.add(key)
                if all(k in current for k in COMBINATION):
                    if main_window.isHidden():
                        self.update.emit(True)
                        self.tray_icon.emit(False)
                    else:
                        self.update.emit(False)
                        self.tray_icon.emit(True)

            if self.entry.hasFocus() and self.result_list.count():
                if key == Key.down:
                    if not self.result_list.currentRow() == (self.result_list.count() - 1):
                        self.current_row.emit(self.parent().nextVisibleItem())

                if key == Key.up:
                    if not self.result_list.currentRow() == 0:
                        self.current_row.emit(self.parent().prevVisibleItem())

            if key == Key.esc:
                listener.stop()

        def on_release(key):
            """ REMOVE """
            try:
                current.remove(key)
            except KeyError:
                pass

        with Listener(on_press=on_press, on_release=on_release) as listener:
            listener.join()


class MainFrame(QFrame):

    # noinspection PyUnresolvedReferences
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("MainFrame")

        _layout = QVBoxLayout()
        _layout.setSpacing(10)
        _layout.setContentsMargins(10, 23, 10, 10)

        self.entry = TextEdit(self)
        self.entry.textChanged.connect(self.textChanged)
        self.highlighter = Highlighter(self.entry.document())

        self.result_list = QListWidget()
        self.result_list.setFixedHeight(0)
        self.result_list.setObjectName("result_list")
        self.result_list.itemDoubleClicked.connect(self.openProgram)
        self.result_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        key_thread = KeyThread(self)
        key_thread.start()

        key_thread.update.connect(self.parent().setVisible)
        key_thread.tray_icon.connect(self.parent().tray_icon.setVisible)
        key_thread.current_row.connect(self.result_list.setCurrentRow)

        _layout.addWidget(self.entry)
        _layout.addWidget(self.result_list)
        _layout.addStretch()
        self.setStyleSheet(readCss("style/main.css"))
        self.setLayout(_layout)
        self.getApps()

    def getResultList(self):
        return self.result_list

    def getEntry(self):
        return self.entry

    def countVisibleItems(self):
        visible_count = 0
        item_count = self.result_list.count()
        for row in range(item_count):
            if not self.result_list.isRowHidden(row):
                visible_count += 1
                if visible_count >= 3:
                    return 3
        return visible_count

    def nextVisibleItem(self):
        for row in range(self.result_list.currentRow() + 1, self.result_list.count()):
            if not self.result_list.isRowHidden(row):
                return row
        return self.result_list.currentRow()

    def prevVisibleItem(self):
        for row in range(self.result_list.currentRow() - 1, -1, -1):
            if not self.result_list.isRowHidden(row):
                return row
        return self.result_list.currentRow()

    def firstVisibleRow(self):
        for row in range(self.result_list.count()):
            if not self.result_list.isRowHidden(row):
                return row
        return -1

    def openProgram(self):
        os.startfile(self.result_list.currentItem().text())
        main_window.hide()

    def textChanged(self):
        text = self.entry.toPlainText()
        if text.isspace() or not text:
            self.result_list.setFixedHeight(0)
            self.parent().setFixedHeight(85)
        else:
            for row in range(self.result_list.count()):
                it = self.result_list.item(row)

                widget = self.result_list.itemWidget(it)
                if text:
                    it.setHidden(not self.filter(text, widget.getText()))
                else:
                    it.setHidden(False)

            _h = self.result_list.sizeHintForRow(0) * self.countVisibleItems()
            self.result_list.setFixedHeight(_h)
            self.parent().setFixedHeight(85 + _h)
            self.result_list.setCurrentRow(self.firstVisibleRow())

    @staticmethod
    def filter(text: str, keywords: str):
        return text.lower() in keywords.lower()

    @staticmethod
    def getAppIcon(path):
        fin = QFileInfo(path)
        model = QFileSystemModel()

        model.setRootPath(fin.path())
        qq = model.iconProvider()
        icon = qq.icon(fin)
        return icon

    def getApps(self):
        for app_path in APPS["APPS"]:
            app_name = str(os.path.basename(app_path)).split(".")[0]

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

        show_action = QAction("Show", self)
        quit_action = QAction("Exit", self)
        hide_action = QAction("Hide", self)
        show_action.triggered.connect(self.show)
        hide_action.triggered.connect(self.hide)
        quit_action.triggered.connect(app.quit)
        tray_menu = QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(quit_action)

        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.setContextMenu(tray_menu)

        _layout = clearLayout(QVBoxLayout())
        self._main_frame = MainFrame(self)
        _layout.addWidget(self._main_frame)

        # Keyboard Shortcuts
        # self.switch_visibility = QShortcut(QKeySequence('Alt+Space'), self)
        # self.switch_visibility.activated.connect(self.switchVisibility)

        # self.frame = QFrame(self)
        # self.frame.setStyleSheet("background: green;")
        # self.frame.setGeometry(0, 0, 250, 5)

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
    program_list = []
    _ALLUSERSPROFILE = os.path.join(os.environ["ALLUSERSPROFILE"], "Start Menu", "Programs")
    _USERPROFILE = os.path.join(os.environ['USERPROFILE'] + r"\AppData\Roaming\Microsoft\Windows\Start Menu")

    for root, dirs, files in os.walk(_ALLUSERSPROFILE):
        for file in files:
            if file.endswith(".lnk"):
                program_list.append(str(os.path.join(root, file)))

    for root, dirs, files in os.walk(_USERPROFILE):
        for file in files:
            if file.endswith(".lnk"):
                program_list.append(str(os.path.join(root, file)))

    with open('database/package.json', 'r', encoding='utf-8') as file:
        JSON = json.load(file)
        JSON["APPS"] = program_list

    with open('database/package.json', 'w', encoding='utf-8') as file:
        json.dump(JSON, file, ensure_ascii=False, indent=4)

    return JSON


if __name__ == "__main__":
    app = QApplication(sys.argv)

    WIDTH = 1000
    HEIGHT = 85
    APPS = None
    SYNTAXS = None

    # TODO: checkout external files
    # TODO: load database / package.json & settings.json
    # TODO: fix resize

    try:
        with open('database/package.json', 'r', encoding='utf-8') as f:
            __JSON = json.load(f)
            if len(__JSON["APPS"]) == 0:
                APPS = getPrograms()
            else:
                APPS = __JSON
            SYNTAXS = __JSON["SYNTAXS"]
    except FileNotFoundError:
        with open('database/package.json', 'w'):
            pass
        APPS = getPrograms()

    # splash_screen = SplashScreen()
    start = time()
    splash = QSplashScreen(QPixmap("image/SplashScreen.png"), Qt.WindowStaysOnTopHint)
    splash.show()
    while time() - start < 1:
        sleep(0.001)
        app.processEvents()
    main_window = MainWindow()
    splash.finish(main_window)

    sys.exit(app.exec_())
