from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QWidget,QFileDialog,QMessageBox
from ExtractDataClass import *


class Ui_Dialog(QWidget):

    def _open_folder_dialog(self, name):
        directory = str(QtWidgets.QFileDialog.getExistingDirectory())
        self.lineEdit_7.setText('{}'.format(directory)) if name == 'input' else self.lineEdit_6.setText(
            '{}'.format(directory))

    def _open_file_dialog(self, name):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        if name == "fault":
            file, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                                    "Excel (*.xls *.xlsx *.xlsm)", options=options)
            self.lineEdit_4.setText('{}'.format(file))
        elif name == "trans":
            file, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                                  "Excel (*.csv)", options=options)
            self.lineEdit_5.setText('{}'.format(file))

    def extract_data(self):
        faults_file_name = self.lineEdit_4.text()
        transaction_file_name = self.lineEdit_5.text()

        input_folder = self.lineEdit_7.text()
        output_folder = self.lineEdit_6.text()

        faults_sheet_name = self.lineEdit.text()
        employees_sheet_name = self.lineEdit_2.text()
        categories_sheet_name = self.lineEdit_3.text()

        transaction_cols = ['עסקה', 'קלדנית',  'תאריך קליטה']

        faults_from_c = self.spinBox.value()
        faults_to_c = self.spinBox_5.value()
        faults_from_r = self.spinBox_7.value()

        emp_from_c = self.spinBox_6.value()
        emp_to_c = self.spinBox_2.value()
        emp_from_r = self.spinBox_8.value()

        cat_from_c = self.spinBox_4.value()
        cat_to_c = self.spinBox_3.value()
        cat_from_r = self.spinBox_9.value()

        from_d = self.dateEdit_2.text()
        to_d = self.dateEdit.text()

        extract = ExtractData(faults_file_name=faults_file_name,
                              transaction_file_name=transaction_file_name,
                              input_folder=input_folder,
                              output_folder=output_folder,
                              faults_sheet_name=faults_sheet_name,
                              employees_sheet_name=employees_sheet_name,
                              categories_sheet_name=categories_sheet_name,
                              transation_cols=transaction_cols,
                              faults_from_c=faults_from_c, faults_to_c=faults_to_c, faults_from_r=faults_from_r,
                              emp_from_c=emp_from_c, emp_to_c=emp_to_c, emp_from_r=emp_from_r,
                              cat_from_c=cat_from_c, cat_to_c=cat_to_c, cat_from_r=cat_from_r,
                              from_d=from_d,to_d=to_d)

        try:
            extract.extract_all()
            extract.load_data()
            QMessageBox.about(self, "Success", "הטעינה בוצעה בהצלחה!")
            QtCore.QCoreApplication.instance().quit()

        except PermissionError as e:
            QMessageBox.about(self, "Failed", "קובץ יצוא פתוח")
        except FileNotFoundError:
            QMessageBox.about(self, "Failed", "קובץ לא נמצא")
        except ValueError:
            QMessageBox.about(self, "Failed", "אין נתונים באחד מהקבצים או יתכן והזנת נתונים שגויים")
        #except Exception:
         #   QMessageBox.about(self, "Failed", "שם גיליון לא קיים/מס' עמודה לא תקין/מס' שורה לא תקין")

    def setupUi(self, dialog):
        dialog.setObjectName("Dialog")
        dialog.resize(446, 356)
        self.buttonBox = QtWidgets.QDialogButtonBox(dialog)
        self.buttonBox.setGeometry(QtCore.QRect(80, 320, 341, 32))
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.gridLayoutWidget = QtWidgets.QWidget(dialog)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 140, 411, 101))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.label_3 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 4, 7, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 3, 7, 1, 1)
        self.label = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 1, 7, 1, 1)
        self.spinBox_5 = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox_5.setObjectName("spinBox_5")
        self.gridLayout.addWidget(self.spinBox_5, 1, 4, 1, 1)
        self.spinBox_3 = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox_3.setObjectName("spinBox_3")
        self.gridLayout.addWidget(self.spinBox_3, 4, 4, 1, 1)
        self.spinBox_6 = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox_6.setObjectName("spinBox_6")
        self.gridLayout.addWidget(self.spinBox_6, 3, 5, 1, 1)
        self.spinBox = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox.setObjectName("spinBox")
        self.gridLayout.addWidget(self.spinBox, 1, 5, 1, 1)
        self.spinBox_4 = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox_4.setObjectName("spinBox_4")
        self.gridLayout.addWidget(self.spinBox_4, 4, 5, 1, 1)
        self.spinBox_2 = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox_2.setObjectName("spinBox_2")
        self.gridLayout.addWidget(self.spinBox_2, 3, 4, 1, 1)
        self.spinBox_8 = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox_8.setObjectName("spinBox_8")
        self.gridLayout.addWidget(self.spinBox_8, 3, 3, 1, 1)
        self.label_13 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_13.setObjectName("label_13")
        self.gridLayout.addWidget(self.label_13, 0, 3, 1, 1)
        self.spinBox_7 = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox_7.setObjectName("spinBox_7")
        self.gridLayout.addWidget(self.spinBox_7, 1, 3, 1, 1)
        self.label_12 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_12.setObjectName("label_12")
        self.gridLayout.addWidget(self.label_12, 0, 4, 1, 1)
        self.label_11 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_11.setObjectName("label_11")
        self.gridLayout.addWidget(self.label_11, 0, 5, 1, 1)
        self.spinBox_9 = QtWidgets.QSpinBox(self.gridLayoutWidget)
        self.spinBox_9.setObjectName("spinBox_9")
        self.gridLayout.addWidget(self.spinBox_9, 4, 3, 1, 1)
        self.label_5 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.label_5, 0, 6, 1, 1)
        self.lineEdit = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.lineEdit.setObjectName("lineEdit")
        self.gridLayout.addWidget(self.lineEdit, 1, 6, 1, 1)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout.addWidget(self.lineEdit_2, 3, 6, 1, 1)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.gridLayout.addWidget(self.lineEdit_3, 4, 6, 1, 1)
        self.gridLayoutWidget_2 = QtWidgets.QWidget(dialog)
        self.gridLayoutWidget_2.setGeometry(QtCore.QRect(10, 20, 411, 108))
        self.gridLayoutWidget_2.setObjectName("gridLayoutWidget_2")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.gridLayoutWidget_2)
        self.gridLayout_3.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.label_10 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_10.setObjectName("label_10")
        self.gridLayout_3.addWidget(self.label_10, 2, 2, 1, 1)
        self.lineEdit_7 = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.lineEdit_7.setReadOnly(True)
        self.gridLayout_3.addWidget(self.lineEdit_7, 2, 1, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_4.setObjectName("label_4")
        self.gridLayout_3.addWidget(self.label_4, 0, 2, 1, 1)
        self.label_8 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_8.setObjectName("label_8")
        self.gridLayout_3.addWidget(self.label_8, 1, 2, 1, 1)
        self.lineEdit_5 = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.lineEdit_5.setReadOnly(True)
        self.gridLayout_3.addWidget(self.lineEdit_5, 1, 1, 1, 1)
        self.lineEdit_4 = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.lineEdit_4.setReadOnly(True)
        self.gridLayout_3.addWidget(self.lineEdit_4, 0, 1, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_9.setObjectName("label_9")
        self.gridLayout_3.addWidget(self.label_9, 3, 2, 1, 1)
        self.lineEdit_6 = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.lineEdit_6.setReadOnly(True)
        self.gridLayout_3.addWidget(self.lineEdit_6, 3, 1, 1, 1)
        self.toolButton_3 = QtWidgets.QToolButton(self.gridLayoutWidget_2)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("./icons/folder.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.toolButton_3.setIcon(icon)
        self.toolButton_3.setObjectName("toolButton_3")
        self.gridLayout_3.addWidget(self.toolButton_3, 2, 0, 1, 1)
        self.toolButton_3.clicked.connect(lambda: self._open_folder_dialog('input'))

        self.toolButton_4 = QtWidgets.QToolButton(self.gridLayoutWidget_2)
        self.toolButton_4.setIcon(icon)
        self.toolButton_4.setObjectName("toolButton_4")
        self.gridLayout_3.addWidget(self.toolButton_4, 3, 0, 1, 1)
        self.toolButton_4.clicked.connect(lambda: self._open_folder_dialog('output'))

        self.toolButton_5 = QtWidgets.QToolButton(self.gridLayoutWidget_2)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("./icons/excel.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.toolButton_5.setIcon(icon1)
        self.toolButton_5.setObjectName("toolButton_5")
        self.gridLayout_3.addWidget(self.toolButton_5, 1, 0, 1, 1)
        self.toolButton_5.clicked.connect(lambda: self._open_file_dialog('trans'))

        self.toolButton_6 = QtWidgets.QToolButton(self.gridLayoutWidget_2)
        self.toolButton_6.setIcon(icon1)
        self.toolButton_6.setObjectName("toolButton_6")
        self.gridLayout_3.addWidget(self.toolButton_6, 0, 0, 1, 1)
        self.toolButton_6.clicked.connect(lambda: self._open_file_dialog('fault'))

        self.retranslateUi(dialog)

        self.gridLayout_3.addWidget(self.toolButton_6, 0, 0, 1, 1)
        self.horizontalLayoutWidget = QtWidgets.QWidget(dialog)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(210, 250, 211, 51))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.dateEdit = QtWidgets.QDateEdit(self.horizontalLayoutWidget)
        self.dateEdit.setObjectName("dateEdit")
        self.horizontalLayout.addWidget(self.dateEdit)
        #self.label_14 = QtWidgets.QLabel(self.horizontalLayoutWidget)
        #self.label_14.setObjectName("label_14")
        #self.horizontalLayout.addWidget(self.label_14)
        self.dateEdit_2 = QtWidgets.QDateEdit(self.horizontalLayoutWidget)
        self.dateEdit_2.setObjectName("dateEdit_2")
        self.horizontalLayout.addWidget(self.dateEdit_2)
        self.label_15 = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label_15.setObjectName("label_15")
        self.horizontalLayout.addWidget(self.label_15)

        # init data
        self.lineEdit_7.setText('C:/Users/Delphine/PycharmProjects/BIDream')
        self.lineEdit_6.setText('C:/Users/Delphine/PycharmProjects/BIDream/output')

        self.lineEdit_4.setText('C:/Users/Delphine/PycharmProjects/BIDream/files/report.xlsm')
        self.lineEdit_5.setText('C:/Users/Delphine/PycharmProjects/BIDream/files/data.csv')

        self.lineEdit.setText('תקלות מוקד')
        self.lineEdit_2.setText('טבלאות לתחזוק')
        self.lineEdit_3.setText('טבלאות לתחזוק')

        self.spinBox.setValue(4)
        self.spinBox_5.setValue(10)
        self.spinBox_7.setValue(5)

        self.spinBox_6.setValue(4)
        self.spinBox_2.setValue(7)
        self.spinBox_8.setValue(2)

        self.spinBox_4.setValue(0)
        self.spinBox_3.setValue(2)
        self.spinBox_9.setValue(2)

        self.dateEdit_2.setDate(QtCore.QDate(2020, 1,1 ))
        self.dateEdit.setDate(QtCore.QDate(2020, 1, 31))
        # Extract data
        self.buttonBox.accepted.connect(self.extract_data)
        self.buttonBox.rejected.connect(dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(dialog)

    def retranslateUi(self, dialog):
        _translate = QtCore.QCoreApplication.translate
        dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label_3.setText(_translate("Dialog", "סגי תקלות"))
        self.label_2.setText(_translate("Dialog", "נציגים"))
        self.label.setText(_translate("Dialog", "תקלות מוקד"))
        self.label_13.setText(_translate("Dialog", "דלג"))
        self.label_12.setText(_translate("Dialog", "עד-"))
        self.label_11.setText(_translate("Dialog", "מ-"))
        self.label_5.setText(_translate("Dialog", "שם גיליון"))
        self.label_4.setText(_translate("Dialog", "קובץ תקלות"))
        self.label_8.setText(_translate("Dialog", "קובץ עסקאות"))
        self.label_9.setText(_translate("Dialog", "תיקיית ייצוא"))
        self.label_10.setText(_translate("Dialog", "תיקיית מקור"))
        #self.label_14.setText(_translate("Dialog", "עד-"))
        #self.label_15.setText(_translate("Dialog", "מ-"))
