import os

import pandas
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QMessageBox, QInputDialog
from main import reset, selectedItem, printEnvelope, printA4, saveItem, deleteAction, clear


class Ui_MainWindow(object):

    def __init__(self):
        self.actionOpen_Excel_File = None
        self.actionOpen_Text_File = None
        self.statusbar = None
        self.menuFile = None
        self.menubar = None
        self.listWidget = None
        self.label = None
        self.line5 = None
        self.line4 = None
        self.line3 = None
        self.line2 = None
        self.line1 = None
        self.topic_comboBox = None
        self.label_2 = None
        self.formLayout = None
        self.checkBox_2 = None
        self.checkBox = None
        self.deleteButton = None
        self.editButton = None
        self.resetButton = None
        self.printButton = None
        self.verticalLayout = None
        self.gridLayout = None
        self.gridLayoutWidget = None
        self.centralwidget = None

    def setupUi(self, main_window):
        main_window.setObjectName("MainWindow")
        main_window.resize(1000, 500)
        self.centralwidget = QtWidgets.QWidget(main_window)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(50, 25, 900, 400))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setHorizontalSpacing(10)
        self.gridLayout.setVerticalSpacing(4)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.printButton = QtWidgets.QPushButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.printButton.setFont(font)
        self.printButton.setObjectName("printButton")
        self.verticalLayout.addWidget(self.printButton)
        self.resetButton = QtWidgets.QPushButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.resetButton.setFont(font)
        self.resetButton.setObjectName("resetButton")
        self.verticalLayout.addWidget(self.resetButton)
        self.editButton = QtWidgets.QPushButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.editButton.setFont(font)
        self.editButton.setObjectName("editButton")
        self.verticalLayout.addWidget(self.editButton)

        self.deleteButton = QtWidgets.QPushButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.deleteButton.setFont(font)
        self.deleteButton.setObjectName("deleteButton")
        self.verticalLayout.addWidget(self.deleteButton)

        self.checkBox = QtWidgets.QCheckBox(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.checkBox.setFont(font)
        self.checkBox.setObjectName("checkBox")
        self.verticalLayout.addWidget(self.checkBox)
        self.checkBox_2 = QtWidgets.QCheckBox(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.checkBox_2.setFont(font)
        self.checkBox_2.setObjectName("checkBox_2")
        self.verticalLayout.addWidget(self.checkBox_2)
        self.gridLayout.addLayout(self.verticalLayout, 0, 2, 1, 1)
        self.formLayout = QtWidgets.QFormLayout()
        self.formLayout.setHorizontalSpacing(4)
        self.formLayout.setVerticalSpacing(10)
        self.formLayout.setObjectName("formLayout")
        self.label_2 = QtWidgets.QLabel(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.ItemRole.LabelRole, self.label_2)

        # name combobox
        self.topic_comboBox = QtWidgets.QComboBox(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.topic_comboBox.setFont(font)
        self.topic_comboBox.setEditable(True)
        self.topic_comboBox.setDuplicatesEnabled(False)
        self.topic_comboBox.completer().setCompletionMode(QtWidgets.QCompleter.CompletionMode.PopupCompletion)
        self.topic_comboBox.completer().setFilterMode(QtCore.Qt.MatchFlag.MatchContains)
        self.topic_comboBox.setObjectName("topic_comboBox")
        self.topic_comboBox.addItems(reset())
        self.topic_comboBox.setCurrentText('')
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.ItemRole.FieldRole, self.topic_comboBox)

        self.line1 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.line1.setFont(font)
        self.line1.setObjectName("line1")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.ItemRole.FieldRole, self.line1)
        self.line2 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.line2.setFont(font)
        self.line2.setObjectName("line2")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.ItemRole.FieldRole, self.line2)
        self.line3 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.line3.setFont(font)
        self.line3.setObjectName("line3")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.ItemRole.FieldRole, self.line3)
        self.line4 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.line4.setFont(font)
        self.line4.setObjectName("line4")
        self.formLayout.setWidget(5, QtWidgets.QFormLayout.ItemRole.FieldRole, self.line4)
        self.line5 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.line5.setFont(font)
        self.line5.setObjectName("line5")
        self.formLayout.setWidget(6, QtWidgets.QFormLayout.ItemRole.FieldRole, self.line5)
        self.label = QtWidgets.QLabel(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.label.setFont(font)
        self.label.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        self.label.setObjectName("label")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.ItemRole.LabelRole, self.label)
        self.gridLayout.addLayout(self.formLayout, 0, 0, 1, 1)

        self.listWidget = QtWidgets.QListWidget(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.listWidget.setFont(font)
        self.listWidget.setObjectName("listWidget")
        self.listWidget.addItems(reset())
        self.gridLayout.addWidget(self.listWidget, 0, 1, 1, 1)

        main_window.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(main_window)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 798, 30))
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.menubar.setFont(font)
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.menuFile.setFont(font)
        self.menuFile.setObjectName("menuFile")
        main_window.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(main_window)
        self.statusbar.setObjectName("statusbar")
        main_window.setStatusBar(self.statusbar)
        self.actionOpen_Text_File = QtGui.QAction(main_window)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.actionOpen_Text_File.setFont(font)
        self.actionOpen_Text_File.setObjectName("actionOpen_Text_File")
        self.actionOpen_Excel_File = QtGui.QAction(main_window)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        self.actionOpen_Excel_File.setFont(font)
        self.actionOpen_Excel_File.setObjectName("actionOpen_Excel_File")
        self.menuFile.addAction(self.actionOpen_Text_File)
        self.menuFile.addAction(self.actionOpen_Excel_File)
        self.menubar.addAction(self.menuFile.menuAction())

        self.retranslateUi(main_window)
        QtCore.QMetaObject.connectSlotsByName(main_window)

        self.resetAll()

        self.actionOpen_Text_File.triggered.connect(lambda: os.startfile('zong.txt'))
        self.actionOpen_Excel_File.triggered.connect(lambda: os.startfile('zong.xlsx'))
        self.topic_comboBox.activated.connect(lambda: self.selectItem(name=self.topic_comboBox.currentText()))
        self.printButton.clicked.connect(self.printAction)
        self.resetButton.clicked.connect(self.resetAll)
        self.editButton.clicked.connect(self.editItem)
        self.deleteButton.clicked.connect(self.deleteItem)
        self.listWidget.itemClicked.connect(lambda: self.selectItem(index=self.listWidget.currentRow()))

    def retranslateUi(self, main_window):
        _translate = QtCore.QCoreApplication.translate
        main_window.setWindowTitle(_translate("MainWindow", "ซอง"))
        self.printButton.setText(_translate("MainWindow", "Print"))
        self.resetButton.setText(_translate("MainWindow", "Reset"))
        self.editButton.setText(_translate("MainWindow", "Edit"))
        self.deleteButton.setText(_translate("MainWindow", "Delete"))
        self.checkBox.setText(_translate("MainWindow", "A4"))
        self.checkBox_2.setText(_translate("MainWindow", "รับเอง"))
        self.label_2.setText(_translate("MainWindow", "เรียน"))
        self.label.setText(_translate("MainWindow", "หัวข้อ"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.actionOpen_Text_File.setText(_translate("MainWindow", "เลขอนุญาต"))
        self.actionOpen_Excel_File.setText(_translate("MainWindow", "ไฟล์ Exel"))

    def resetAll(self):
        self.topic_comboBox.clear()
        self.topic_comboBox.addItems(reset())
        self.topic_comboBox.setCurrentText('')
        self.topic_comboBox.setDisabled(False)
        self.listWidget.clear()
        self.listWidget.addItems(reset())
        self.line1.clear()
        self.line1.setDisabled(True)
        self.line2.clear()
        self.line2.setDisabled(True)
        self.line3.clear()
        self.line3.setDisabled(True)
        self.line4.clear()
        self.line4.setDisabled(True)
        self.line5.clear()
        self.line5.setDisabled(True)
        self.printButton.setDisabled(False)
        self.deleteButton.setDisabled(False)
        self.editButton.setText('Edit')

    def printAction(self):
        des_list = reset()
        if self.topic_comboBox.currentText() in des_list:
            if self.topic_comboBox.currentText() == 'เพิ่มที่อยู่':
                return None
            name = None
            if self.checkBox_2.isChecked() is True:
                name = self.printBox()
                if name is None:
                    return None
            if self.checkBox.isChecked():
                printA4(self.topic_comboBox.currentText(), self.checkBox_2.isChecked(), name)
            else:
                printEnvelope(self.topic_comboBox.currentText(), self.checkBox_2.isChecked(), name)

    def editItem(self):
        newName = []
        nameList = reset()
        if self.editButton.text() == 'Edit':
            if self.topic_comboBox.currentText() in nameList:
                self.printButton.setDisabled(True)
                self.deleteButton.setDisabled(True)
                if self.topic_comboBox.currentText() != 'เพิ่มที่อยู่':
                    self.topic_comboBox.setDisabled(True)
                self.line1.setDisabled(False)
                self.line2.setDisabled(False)
                self.line3.setDisabled(False)
                self.line4.setDisabled(False)
                self.line5.setDisabled(False)
                self.editButton.setText('Save')
        else:
            newName.append(self.topic_comboBox.currentText())
            newName.append(self.line1.text())
            newName.append(self.line2.text())
            newName.append(self.line3.text())
            newName.append(self.line4.text())
            newName.append(self.line5.text())
            if self.topic_comboBox.currentText() != 'เพิ่มที่อยู่':
                if self.topic_comboBox.currentText() in nameList:
                    currentIndex = nameList.index(self.topic_comboBox.currentText())
                    saveItem(currentIndex, newName)
                else:
                    saveItem(-1, newName)
                self.resetAll()

    def deleteItem(self):
        nameList = reset()
        if self.topic_comboBox.currentText() in nameList:
            currentIndex = nameList.index(self.topic_comboBox.currentText())
            if self.topic_comboBox.currentText() == 'เพิ่มที่อยู่' or self.topic_comboBox.currentText() == nameList[0]:
                return None
            elif self.deleteBox():
                deleteAction(currentIndex)
                self.resetAll()

    def selectItem(self, index=None, name=None):
        self.listWidget.clear()
        self.listWidget.addItems(reset())
        self.line1.clear()
        self.line1.setDisabled(True)
        self.line2.clear()
        self.line2.setDisabled(True)
        self.line3.clear()
        self.line3.setDisabled(True)
        self.line4.clear()
        self.line4.setDisabled(True)
        self.line5.clear()
        self.line5.setDisabled(True)
        self.editButton.setText('Edit')

        name_list = reset()
        data_list = []
        if index is not None:
            data_list = selectedItem(name_list[index])
        if name is not None:
            data_list = selectedItem(name)
        for i in range(len(data_list)):
            if i == 0:
                self.topic_comboBox.setCurrentText(data_list[i])
            elif not pandas.isna(data_list[i]):
                exec(f'self.line{i}.setText("{data_list[i]}")')

    @staticmethod
    def printBox():
        msg = QInputDialog()
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        msg.setFont(font)
        msg.setWindowTitle('ชื่อผู้รับ')
        x = msg.exec()
        if x:
            return msg.textValue()
        else:
            return None

    def deleteBox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Warning)
        font = QtGui.QFont()
        font.setFamily("TH Sarabun New")
        font.setPointSize(18)
        msg.setFont(font)
        msg.setWindowTitle('Delete')
        msg.setText('Delete "' + self.topic_comboBox.currentText() + '"')
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        x = msg.exec()

        if x == QMessageBox.StandardButton.Yes:
            return True
        else:
            return False


if __name__ == "__main__":
    import sys

    clear()
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
