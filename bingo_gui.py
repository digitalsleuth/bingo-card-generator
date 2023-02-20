#!/usr/bin/env python3
'''
Initiate the GUI and execute the application components
from within bingo_card_generator
'''

import os
from PyQt6 import QtGui, QtWidgets
import bingo_card_generator

basedir = os.path.dirname(__file__)
description = bingo_card_generator.__description__

try:
    from ctypes import windll
    AppId = 'digitalsleuth.interactive-bingo-card-generator.gui.v6-0-1'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(AppId)
except ImportError:
    pass

class BingoCard(QtWidgets.QMainWindow, bingo_card_generator.UiDialog):
    """BingoCard Class"""
    def __init__(self, parent=None):
        """Call and setup the UI"""
        super(BingoCard, self).__init__(parent)
        self.setup_ui(self)

def main():
    """Execute the application"""
    bingo_app = QtWidgets.QApplication([description, 'windows:darkmode=2'])
    bingo_app.setWindowIcon(QtGui.QIcon(os.path.join(basedir, 'bingo.ico')))
    bingo_app.setStyle('Fusion')
    bingo_form = BingoCard()
    bingo_form.show()
    bingo_app.exec()

if __name__ == '__main__':
    main()
