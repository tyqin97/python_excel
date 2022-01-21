import sys
from tkinter import W
from PyQt6.uic import loadUi
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
import time
from PyQt6.QtWidgets import QWidget
from PyQt6 import QtGui
import os

from Logger import Logger
logger = Logger.Log()


class SplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('DMT5 File Generator')
        self.setFixedSize(700, 450)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self.counter = 0
        self.n = 300  # total instance

        self.initUI()

        self.timer = QTimer()
        self.timer.timeout.connect(self.loading)
        self.timer.start(30)

    def initUI(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.frame = QFrame()
        self.frame.setObjectName('SplashFrame')
        layout.addWidget(self.frame)

        self.labelTitle = QLabel(self.frame)
        self.labelTitle.setObjectName('LabelTitle')

        # center labels
        self.labelTitle.resize(self.width() - 10, 130)
        self.labelTitle.move(0, 40)  # x, y
        self.labelTitle.setText('DMT5 File Generator')
        self.labelTitle.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.labelDescription = QLabel(self.frame)
        self.labelDescription.resize(self.width() - 10, 50)
        self.labelDescription.move(0, self.labelTitle.height())
        self.labelDescription.setObjectName('LabelDesc')
        self.labelDescription.setText('Initializing...')
        self.labelDescription.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.progressBar = QProgressBar(self.frame)
        self.progressBar.setObjectName('SplashProgress')
        self.progressBar.resize(self.width() - 200 - 10, 30)
        self.progressBar.move(100, self.labelDescription.y() + 130)
        self.progressBar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progressBar.setFormat('%p%')
        self.progressBar.setTextVisible(True)
        self.progressBar.setRange(0, self.n)
        self.progressBar.setValue(20)

        self.labelLoading = QLabel(self.frame)
        self.labelLoading.resize(self.width() - 10, 50)
        self.labelLoading.move(0, self.progressBar.y() + 70)
        self.labelLoading.setObjectName('LabelLoading')
        self.labelLoading.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.labelLoading.setText('Loading...')

    def loading(self):
        self.progressBar.setValue(self.counter)

        if self.counter == int(self.n * 0.3):
            self.labelDescription.setText(
                'Setting Things Up...')
        elif self.counter == int(self.n * 0.8):
            self.labelDescription.setText(
                'Almost Done...')
        elif self.counter >= self.n:
            self.timer.stop()
            self.close()

            time.sleep(1)

            self.myApp = MainWindow()
            self.myApp.show()

        self.counter += 1


class Worker(QRunnable):

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()

        self.fn = fn
        self.args = args
        self.kwargs = kwargs

    @pyqtSlot()
    def run(self):
        self.fn(*self.args, **self.kwargs)
        splash.myApp.doneSignal.finished.emit()


class Signals(QObject):
    finished = pyqtSignal()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.comboText = "---"
        self.threadPool = QThreadPool()
        self.doneSignal = Signals()
        self.errorSignal = Signals()
        self.succSignal = Signals()
        self.genSignal = Signals()
        self.cwd = os.getcwd()
        self.loadUI()
        self.loadDMT()

        self.show()

    def confirmationBox(self):
        msgBox = QMessageBox()
        msgBox.setWindowTitle("Confirmation")
        msgBox.setText("Are You Sure?")
        msgBox.setIcon(QMessageBox.Icon.Question)
        msgBox.setStandardButtons(
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msgBox.setDefaultButton(QMessageBox.StandardButton.Yes)
        msgBox.setWindowIcon(QtGui.QIcon('./img/msgBox.ico'))
        self.confirmationTxt = msgBox.exec()

    def loadUI(self):
        loadUi("MainWindow.ui", self)
        self.browseButton.clicked.connect(self.browser)
        self.importButton.clicked.connect(self.startImport)
        self.confirmButton.clicked.connect(self.startGenerate)
        self.comboBox.currentTextChanged.connect(self.onChanged)

        self.genWidget.setEnabled(False)

    def loadDMT(self):
        from Main import DMT5
        self.mainModule = DMT5()

    def browser(self):
        self.fname = QFileDialog.getOpenFileName(
            self, 'Open File', 'C:\\')
        self.pathLine.setText(self.fname[0])

    def getImportStatus(self):
        self.progressBarInc(80, 100)

    def onChanged(self, value):
        self.comboText = value

    def startGenerate(self):
        if self.comboText == "---":
            self.checkGenSelection()
        elif self.comboText != None:
            try:
                self.mainModule.selectGenFile(self.comboText.upper())
            except Exception as err:
                logger.error(err)
            finally:
                self.doneGenerateBox()

    def startImport(self):
        self.confirmationBox()

        if QMessageBox.StandardButton(self.confirmationTxt) is QMessageBox.StandardButton.Yes:
            self.progressBarInc(0, 80)

            self.worker2 = Worker(self.loadExcel)
            self.threadPool.start(self.worker2)

            self.doneSignal.finished.connect(self.getImportStatus)

            self.errorSignal.finished.connect(self.importErrorBox)

            self.succSignal.finished.connect(self.importSuccessBox)
        else:
            pass

    def progressBarInc(self, s_val, val):
        progValue = s_val
        while progValue <= val:
            self.progressBar.setValue(progValue)
            time.sleep(0.01)
            progValue += 1

    def checkGenSelection(self):
        msgBox = QMessageBox()
        msgBox.setWindowTitle("Generate Error")
        msgBox.setText("Please Select Valid Option To Generate")
        msgBox.setIcon(QMessageBox.Icon.Warning)
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setDefaultButton(QMessageBox.StandardButton.Ok)
        msgBox.setWindowIcon(QtGui.QIcon('./img/msgBox.ico'))
        x = msgBox.exec()

    def doneGenerateBox(self):
        self.setString = "File Successfully Saved In: " + \
            str(self.cwd) + "\\files\\output\\" + \
            str(self.mainModule.timestamp)
        self.pathString = str(self.cwd) + "\\files\\output\\" + \
            str(self.mainModule.timestamp)
        msgBox = QMessageBox()
        msgBox.setWindowTitle("Generate Successful")
        msgBox.setText(self.setString)
        msgBox.setIcon(QMessageBox.Icon.Information)
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setWindowIcon(QtGui.QIcon('./img/msgBox.ico'))
        copyBtn = msgBox.addButton(
            'Copy Path', QMessageBox.ButtonRole.ActionRole)
        copyBtn.clicked.disconnect()
        copyBtn.clicked.connect(self.addToClipBoard)
        msgBox.exec()

    def addToClipBoard(self):
        command = 'echo ' + self.pathString + '| clip'
        os.system(command)

    def importErrorBox(self):
        msgBox = QMessageBox()
        msgBox.setWindowTitle("Import Status")
        msgBox.setText(
            "Import UnSuccessfully \n Please Go To './errorLog/{today_date}.log' \n For More Information")
        msgBox.setIcon(QMessageBox.Icon.Warning)
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setDefaultButton(QMessageBox.StandardButton.Ok)
        msgBox.setWindowIcon(QtGui.QIcon('./img/msgBox.ico'))
        x = msgBox.exec()

    def importSuccessBox(self):
        msgBox = QMessageBox()
        msgBox.setWindowTitle("Import Status")
        msgBox.setText("Import Successfully")
        msgBox.setIcon(QMessageBox.Icon.Information)
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setDefaultButton(QMessageBox.StandardButton.Ok)
        msgBox.setWindowIcon(QtGui.QIcon('./img/msgBox.ico'))
        xxx = msgBox.exec()

        if QMessageBox.StandardButton(xxx) is QMessageBox.StandardButton.Ok:
            self.showUserInput()

    def loadExcel(self):
        self.pathText = self.pathLine.text()
        try:
            self.mainModule.loadWorkbook(path=self.pathText)
        except Exception as err:
            logger.error(err)
            self.errorSignal.finished.emit()
        else:
            self.succSignal.finished.emit()

    def showUserInput(self):
        userString = "Total Module: " + str(self.mainModule.count_sheet_wb_main) + "\nDrawing Number: " + str(self.mainModule.main_drawing_number) + \
                     "\nECO Group ID: " + str(self.mainModule.main_eco_group_id) + "\nEffective Date: " + str(self.mainModule.main_effective_date) + \
                     "\nRead Until: " + str(self.mainModule.main_total_module) + \
            "\nTotal Fab Part: " + str(self.mainModule.main_total_fab_part)
        msgBox = QMessageBox()
        msgBox.setWindowTitle("Display Info")
        msgBox.setText(userString)
        msgBox.setIcon(QMessageBox.Icon.Information)
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setDefaultButton(QMessageBox.StandardButton.Ok)
        msgBox.setWindowIcon(QtGui.QIcon('./img/msgBox.ico'))
        msgBox.exec()

        self.genWidget.setEnabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("./img/main.ico"))
    app.setStyleSheet('''
        #LabelTitle {
            font-size: 60px;
            color: #93deed;
        }

        #LabelDesc {
            font-size: 30px;
            color: #c2ced1;
        }

        #LabelLoading {
            font-size: 30px;
            color: #e8e8eb;
        }

        #SplashFrame {
            background-color: #2F4454;
            color: rgb(220, 220, 220);
        }

        #SplashProgress {
            background-color: #DA7B93;
            color: rgb(200, 200, 200);
            border-style: none;
            border-radius: 10px;
            text-align: center;
            font-size: 30px;
        }

        #SplashProgress::chunk {
            border-radius: 10px;
            background-color: qlineargradient(spread:pad x1:0, x2:1, y1:0.511364, y2:0.523, stop:0 #1C3334, stop:1 #376E6F);
        }
    ''')
    splash = SplashScreen()
    splash.show()

    try:
        sys.exit(app.exec())
    except SystemExit:
        print('Closing Window...')
