#!/usr/bin/env python3
from PyQt5 import QtCore, QtGui, QtWidgets
import bingo_card_generator
import os

basedir = os.path.dirname(__file__)

try:
    from ctypes import windll
    AppId = 'digitalsleuth.interactive-bingo-card-generator.gui.v5-0-0'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(AppId)
except ImportError:
    pass

class BingoCard(QtWidgets.QMainWindow, bingo_card_generator.Ui_Dialog):
    def __init__(self, parent=None):
        super(BingoCard, self).__init__(parent)
        self.setupUi(self)

def main():
    app = QtWidgets.QApplication([])
    #app.setWindowIcon(QtGui.QIcon(os.path.join(basedir, 'bingo.ico')))
    form = BingoCard()
    form.show()
    app.exec_()

if __name__ == '__main__':
    main()
