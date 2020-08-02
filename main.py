from PyQt5 import QtWidgets
from Ui_DialogClass import Ui_Dialog

if __name__ == "__main__":

    import sys

    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
