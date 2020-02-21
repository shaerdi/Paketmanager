# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'LeistungswahldialogUI.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(322, 383)
        self.verticalLayout = QtWidgets.QVBoxLayout(Dialog)
        self.verticalLayout.setObjectName("verticalLayout")
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        spacerItem = QtWidgets.QSpacerItem(30, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout.addItem(spacerItem, 0, 1, 1, 1)
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 3, 2, 1, 1)
        self.NeueLeistung = QtWidgets.QLineEdit(Dialog)
        self.NeueLeistung.setObjectName("NeueLeistung")
        self.gridLayout.addWidget(self.NeueLeistung, 0, 2, 1, 1)
        self.line = QtWidgets.QFrame(Dialog)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.gridLayout.addWidget(self.line, 1, 0, 2, 3)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setObjectName("label_3")
        self.verticalLayout_2.addWidget(self.label_3)
        self.radioButton_UND = QtWidgets.QRadioButton(Dialog)
        self.radioButton_UND.setObjectName("radioButton_UND")
        self.verticalLayout_2.addWidget(self.radioButton_UND)
        self.radioButton_ODER = QtWidgets.QRadioButton(Dialog)
        self.radioButton_ODER.setObjectName("radioButton_ODER")
        self.verticalLayout_2.addWidget(self.radioButton_ODER)
        self.radioButton_NICHT = QtWidgets.QRadioButton(Dialog)
        self.radioButton_NICHT.setObjectName("radioButton_NICHT")
        self.verticalLayout_2.addWidget(self.radioButton_NICHT)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem1)
        self.gridLayout.addLayout(self.verticalLayout_2, 4, 0, 1, 1)
        self.listView_Vorschlaege = QtWidgets.QListView(Dialog)
        self.listView_Vorschlaege.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listView_Vorschlaege.setProperty("showDropIndicator", False)
        self.listView_Vorschlaege.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.listView_Vorschlaege.setObjectName("listView_Vorschlaege")
        self.gridLayout.addWidget(self.listView_Vorschlaege, 4, 2, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.verticalLayout.addWidget(self.buttonBox)

        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label_2.setText(_translate("Dialog", "Vorhanden im aktuellen Excel"))
        self.label.setText(_translate("Dialog", "Neue Leistung"))
        self.label_3.setText(_translate("Dialog", "Hinzufügen zu"))
        self.radioButton_UND.setText(_translate("Dialog", "UND"))
        self.radioButton_ODER.setText(_translate("Dialog", "ODER"))
        self.radioButton_NICHT.setText(_translate("Dialog", "NICHT"))

