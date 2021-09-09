#!/usr/bin/env python3
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication
import bingo_card_generator

class BingoCard(QtWidgets.QMainWindow, bingo_card_generator.Ui_Dialog):
    def __init__(self, parent=None):
        super(BingoCard, self).__init__(parent)
        self.setupUi(self)
        #self.setWindowIcon(QtGui.QIcon('bingo.ico'))

def main():
    app = QApplication([])
    form = BingoCard()
    form.show()
    app.exec_()

if __name__ == '__main__':
    main()
