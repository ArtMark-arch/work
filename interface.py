# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'interface.ui'
#
# Created by: PyQt5 UI code generator 5.15.2
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(342, 165)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.pick_btn = QtWidgets.QPushButton(self.centralwidget)
        self.pick_btn.setObjectName("pick_btn")
        self.gridLayout.addWidget(self.pick_btn, 1, 0, 1, 1)
        self.search_btn = QtWidgets.QPushButton(self.centralwidget)
        self.search_btn.setObjectName("search_btn")
        self.gridLayout.addWidget(self.search_btn, 1, 1, 1, 1)
        self.data_input = QtWidgets.QLineEdit(self.centralwidget)
        self.data_input.setObjectName("data_input")
        self.gridLayout.addWidget(self.data_input, 0, 0, 1, 2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setStatusTip("")
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "поиск товара"))
        self.pick_btn.setText(_translate("MainWindow", "выбрать файл"))
        self.search_btn.setText(_translate("MainWindow", "поиск"))