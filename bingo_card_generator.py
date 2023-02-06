#!/usr/bin/env python3

'''
Tool to generate a clickable bingo card for virtual bingo events and to
generate PDF's for printing these cards.

Requires pdfkit and wkhtmltopdf for OS it's being run on
(Windows requires an exe, linux install from pkg manager)

This script currently will only generate 6 cards on one sheet (2 rows of 3 cards).
Hex codes can be used in place of words for colours,
but must be wrapped in quotes on the command line

While not normally necessary for CSS, the float values are required for
the proper printing of the PDF's with wkhtmltopdf.
If these are removed, you can expect the PDF's to be off-center or misaligned.

CSS Maple Leaf author Andre Lopes - https://codepen.io/alldrops/pen/jAzZmw
'''

from PyQt5 import QtCore, QtGui, QtWidgets
import argparse
from random import Random
import re
import base64
import csv
import imghdr
import os
import pdfkit
import sys
from PIL import Image
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, NamedStyle, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

__author__ = 'Corey Forman'
__date__ = '05 Feb 2023'
__version__ = '4.0.0'
__description__ = 'Interactive Bingo Card and PDF Generator'
__colour_groups__ = 'https://www.w3schools.com/colors/colors_groups.asp'


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Palette = QtGui.QPalette()
        Palette.setColor(Palette.ColorRole.Window, QtGui.QColor('#DDDDDD'))
        Dialog.setObjectName("Dialog")
        Dialog.setFixedSize(400, 300)
        Dialog.setPalette(Palette)
        self.allow_select = QtWidgets.QCheckBox("",self)
        self.allow_select.setObjectName("allow_select")
        self.allow_select.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.allow_select.stateChanged.connect(self.selectBox)
        self.allow_select.setGeometry(12, 235, 178, 40)
        self.allow_select.setStyleSheet("QCheckBox::indicator {width: 18px; height: 18px;}")
        self.allow_select_label = QtWidgets.QLabel(Dialog)
        self.allow_select_label.setGeometry(QtCore.QRect(12, 235, 140, 40))
        self.allow_select_label.setToolTipDuration(-1)
        self.allow_select_label.setObjectName("allow_select_label")
        self.dauber_shape = QtWidgets.QComboBox(Dialog)
        self.dauber_shape.setGeometry(QtCore.QRect(140, 167, 131, 27))
        self.dauber_shape.setEditable(False)
        self.dauber_shape.setObjectName("dauber_shape")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.addItem("")
        self.dauber_shape.currentIndexChanged[int].connect(self.on_Select)
        self.dauber_shape_label = QtWidgets.QLabel(Dialog)
        self.dauber_shape_label.setGeometry(QtCore.QRect(10, 167, 121, 21))
        self.dauber_shape_label.setToolTipDuration(-1)
        self.dauber_shape_label.setObjectName("dauber_shape_label")
        self.title_label = QtWidgets.QLabel(Dialog)
        self.title_label.setGeometry(QtCore.QRect(4, 2, 391, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.title_label.setFont(font)
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.title_label.setObjectName("title_label")
        self.dauber_colour_label = QtWidgets.QLabel(Dialog)
        self.dauber_colour_label.setGeometry(QtCore.QRect(10, 132, 121, 21))
        self.dauber_colour_label.setToolTipDuration(-1)
        self.dauber_colour_label.setObjectName("dauber_colour_label")
        self.card_colour_label = QtWidgets.QLabel(Dialog)
        self.card_colour_label.setGeometry(QtCore.QRect(10, 92, 121, 21))
        self.card_colour_label.setToolTipDuration(-1)
        self.card_colour_label.setObjectName("card_colour_label")
        self.select_logo = QtWidgets.QPushButton(Dialog)
        self.select_logo.setGeometry(QtCore.QRect(290, 167, 100, 27))
        self.select_logo.setDefault(False)
        self.select_logo.setObjectName("select_logo")
        self.select_logo.clicked.connect(self.selectLogo)
        self.select_logo.setEnabled(False)
        self.select_logo.setVisible(False)
        self.select_result = QtWidgets.QLineEdit(Dialog)
        self.select_result.setObjectName("select_result")
        self.select_result.setReadOnly(True)
        self.select_result.setVisible(False)
        self.select_result.setGeometry(QtCore.QRect(9, 205, 382, 30))
        self.generate = QtWidgets.QPushButton(Dialog)
        self.generate.setGeometry(QtCore.QRect(290, 47, 100, 31))
        self.generate.setDefault(False)
        self.generate.setObjectName("generate")
        self.generate.clicked.connect(lambda: guiEverything(int(self.number.text()),self.card_colour.text(), self.dauber_colour.text(), self.dauber_shape.currentText(), self.getDirectory(), self.selectLogo(), self.selectBox()))
        self.close = QtWidgets.QPushButton(Dialog)
        self.close.setGeometry(QtCore.QRect(290, 87, 100, 31))
        self.close.setObjectName("close")
        self.close.clicked.connect(QtWidgets.QApplication.instance().quit)
        self.number_label = QtWidgets.QLabel(Dialog)
        self.number_label.setGeometry(QtCore.QRect(10, 52, 121, 21))
        self.number_label.setToolTipDuration(-1)
        self.number_label.setObjectName("number_label")
        self.number = QtWidgets.QLineEdit(Dialog)
        self.number.setGeometry(QtCore.QRect(140, 47, 131, 31))
        self.number.setPlaceholderText("")
        self.number.setObjectName("number")
        self.card_colour = QtWidgets.QLineEdit(Dialog)
        self.card_colour.setGeometry(QtCore.QRect(140, 87, 131, 31))
        self.card_colour.setObjectName("card_colour")
        self.card_colour_picker = QtWidgets.QPushButton(Dialog)
        self.card_colour_picker.setObjectName("card_colour_picker")
        self.card_colour_picker.clicked.connect(self.cardColourPicker)
        self.card_colour_picker.setGeometry(QtCore.QRect(110, 92, 20, 20))
        self.card_colour_picker.setStyleSheet("background-color: blue; border: 1px solid black")
        self.dauber_colour = QtWidgets.QLineEdit(Dialog)
        self.dauber_colour.setGeometry(QtCore.QRect(140, 127, 131, 31))
        self.dauber_colour.setObjectName("dauber_colour")
        self.dauber_colour.setEnabled(False)
        self.dauber_colour_picker = QtWidgets.QPushButton(Dialog)
        self.dauber_colour_picker.setObjectName("dauber_colour_picker")
        self.dauber_colour_picker.clicked.connect(self.dauberColourPicker)
        self.dauber_colour_picker.setGeometry(QtCore.QRect(110, 132, 20, 20))
        self.dauber_colour_picker.setStyleSheet("background-color: red; border: 1px solid black")
        self.dauber_colour_picker.setEnabled(False)
        self.version_label = QtWidgets.QLabel(Dialog)
        self.version_label.setGeometry(QtCore.QRect(4, 19, 391, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.allow_select.setFont(font)
        self.allow_select_label.setFont(font)
        self.version_label.setFont(font)
        self.version_label.setAlignment(QtCore.Qt.AlignCenter)
        self.version_label.setObjectName("version_label")
        self.dauber_shape_label.setBuddy(self.dauber_shape)
        self.dauber_colour_label.setBuddy(self.dauber_colour)
        self.card_colour_label.setBuddy(self.card_colour)
        self.number_label.setBuddy(self.number)
        self.allow_select_label.setBuddy(self.allow_select)
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        Dialog.setTabOrder(self.number, self.card_colour)
        Dialog.setTabOrder(self.card_colour, self.dauber_colour)
        Dialog.setTabOrder(self.dauber_colour, self.dauber_shape)
        Dialog.setTabOrder(self.dauber_shape, self.select_logo)
        Dialog.setTabOrder(self.select_logo, self.allow_select)
        Dialog.setTabOrder(self.allow_select, self.generate)
        Dialog.setTabOrder(self.generate, self.close)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Interactive Bingo Card Generator"))
        self.dauber_shape.setCurrentText(_translate("Dialog", "Circle"))
        self.dauber_shape.setItemText(0, _translate("Dialog", "Checkmark"))
        self.dauber_shape.setItemText(1, _translate("Dialog", "Circle"))
        self.dauber_shape.setItemText(2, _translate("Dialog", "Clover"))
        self.dauber_shape.setItemText(3, _translate("Dialog", "Heart"))
        self.dauber_shape.setItemText(4, _translate("Dialog", "Logo"))
        self.dauber_shape.setItemText(5, _translate("Dialog", "Maple-Leaf"))
        self.dauber_shape.setItemText(6, _translate("Dialog", "Moon"))
        self.dauber_shape.setItemText(7, _translate("Dialog", "Square"))
        self.dauber_shape.setItemText(8, _translate("Dialog", "Star"))
        self.dauber_shape.setItemText(9, _translate("Dialog", "Unicorn"))
        self.dauber_shape.setItemText(10, _translate("Dialog", "X-Mark"))
        self.dauber_shape_label.setToolTip(_translate("Dialog", "Choose the shape of the dauber"))
        self.dauber_shape_label.setText(_translate("Dialog", "Dauber Shape"))
        self.title_label.setText(_translate("Dialog", "Interactive Bingo Card and PDF Generator"))
        self.dauber_colour.setText(_translate("Dialog", "Red"))
        self.dauber_colour_label.setToolTip(_translate("Dialog", "Choose the colour of the dauber"))
        self.dauber_colour_label.setText(_translate("Dialog", "Dauber Colour"))
        self.card_colour.setText(_translate("Dialog", "Blue"))
        self.card_colour_label.setToolTip(_translate("Dialog", "Choose the colour of the card"))
        self.card_colour_label.setText(_translate("Dialog", "Card Colour"))
        self.generate.setText(_translate("Dialog", "Generate"))
        self.select_logo.setText(_translate("Dialog", "Select Logo"))
        self.close.setText(_translate("Dialog", "Close"))
        self.number_label.setToolTip(_translate("Dialog", "Choose the number of cards"))
        self.number_label.setText(_translate("Dialog", "# of Cards"))
        self.allow_select.setText(_translate("Dialog", "Allow dauber selection?"))
        self.allow_select_label.setToolTip(_translate("Dialog", "Allow the player to change their dauber after generation of card"))
        self.card_colour.setPlaceholderText(_translate("Dialog", "Blue"))
        self.dauber_colour.setPlaceholderText(_translate("Dialog", "Red"))
        self.version_label.setText(_translate("Dialog", __version__ + " - " + __date__))

    def getDirectory(self):
        button = QtWidgets.QFileDialog()
        button.setFileMode(QtWidgets.QFileDialog.Directory)
        button.setOption(QtWidgets.QFileDialog.ShowDirsOnly)
        chosenPath = button.getExistingDirectory(self, 'Select the output location for your cards ...', os.path.curdir)

        return chosenPath

    def selectLogo(self):
        index = self.dauber_shape.currentIndex()
        if ((index == 4) and self.select_result.text() == ''):
            selected_file, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select the image you wish to use...","", "Image Files (*.jpg *.png)")
            self.select_result.setText(selected_file)
        elif ((index == 4) and self.select_result.text() != ''):
            selected_file = self.select_result.text()
        else:
            selected_file = False

        return selected_file

    def selectBox(self):
        if self.allow_select.isChecked():
            return True
        else:
            return False

    def on_Select(self, index):
        self.select_logo.setEnabled(index == 4)
        self.select_logo.setVisible(index == 4)
        self.select_result.setVisible(index == 4)
        self.dauber_colour.setEnabled(index in (1, 3, 5, 7))
        self.dauber_colour_picker.setEnabled(index in (1, 3, 5, 7))

    def dauberColourPicker(self):
        dauberColour = QtWidgets.QColorDialog.getColor()
        if dauberColour.isValid():
            self.dauber_colour_picker.setStyleSheet('QPushButton { background-color: ' + dauberColour.name() + '; border: 1px solid black}')
            self.dauber_colour.setText(dauberColour.name())
        else:
            self.dauber_colour_picker.setStyleSheet('QPushButton { background-color: red; border: 1px solid black}')
            self.dauber_colour.setText('Red')

    def cardColourPicker(self):
        cardColour = QtWidgets.QColorDialog.getColor()
        if cardColour.isValid():
            self.card_colour_picker.setStyleSheet('QPushButton { background-color: ' + cardColour.name() + '; border: 1px solid black}')
            self.card_colour.setText(cardColour.name())
        else:
            self.card_colour_picker.setStyleSheet('QPushButton { background-color: blue; border: 1px solid black}')
            self.card_colour.setText('Blue')

def guiEverything(number, card_colour, dauber_colour, dauber_shape, output, logo, allow_select):
    args = {'num': number, 'pdf': True, 'card_colour': card_colour, 'dauber_colour': dauber_colour, 'dauber_shape': dauber_shape, 'logo': logo, 'allow_select': allow_select, 'base_colour': card_colour, 'output': output, 'excel': str((card_colour).strip("#") + '-cards.xlsx'), 'everything': True}
    createCard(args)
    grabNumbers(args)
    writeToExcel(args['num'], args['card_colour'], args['excel'], args['output'])
    msgBox = QtWidgets.QMessageBox()
    msgBox.setWindowTitle("Finished")
    msgBox.setText("All files created in {}\n\n{} {} card(s) created with a {}, {} dauber.\n\n{} Excel file also created for tracking called numbers.\n\nYou may close the Bingo Card Generator now, or generate more cards if you wish.".format(str(args['output']), str(args['num']), args['card_colour'], args['dauber_colour'], args['dauber_shape'], args['excel']))
    msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msgBox.exec_()


def createCard(arguments):
    """Creates the HTML version of the card"""
    card_colour = arguments['card_colour'].lower()
    dauber_colour = arguments['dauber_colour'].lower()
    dauber_shape = arguments['dauber_shape'].lower()
    output_path = arguments['output'] + os.sep
    if not os.path.exists(output_path):
        os.mkdir(output_path)
    total = 1
    card_title = ""
    open_head = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
'''
    open_style = "<style>"
    if arguments['logo'] == False or arguments['logo'] == '':
        selected_logo = 'url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACgAAAAyCAYAAAAus5mQAAAABmJLR0QAAABxAK/qGQdhAAAACXBIWXMAAA7DAAAOwwHHb6hkAAAR4ElEQVRYw81ZaZgU1bl+q7qqeqnee3p69n0GZoMBhh0ElEUGUUEMKBhQk5gYDSEGvUluNpN4Ta5LNKKicQkkcQkgi8ggkgEHGGZhmH3fp6ent+l9ra4lP7iPXh/jFaLeJ+d31Tnveb/3/c53zkfgSxxN9ecJU6xZMz4RCF1350/EL2NO8ssCNzk8RrCJ8WVC0H0qLT/v5mNvvUv8WwH0hhMAFygNh8NJBBdOo+Vy/FsA7O7uIAGgpCxfCoQ9o35BnjnS3+G68dZVEgBcvnCG6u9q+pfZ/Jd+rL94QU6I0TyDTn99wOeamZyec2JybLQpNNawNzHesZbXpdVYKrds1bKy+bFwYJ2KZds5gawW4+Jo2dz58a8U4PvVhxiLOW0ZN1n/uIJ39xMy2hPjyPlhOn1ccvXO1FGhnCC0g6Iur0lDuPMYmrmk1DCacJwpoIzlj3ntk9VLqjZyX1mIjVqzWfL2/EwTHnh90CZtLVv/s/tJkVivnzrbKBet9oRAgvB2ezTO2g4jI6sqr/rBd9Rs+XZWkXgtMt70kDE9K/krZbCx9uRdlKvuoZF+dhVz4E1KlpO9XeaYnO81GP5GVqYnp0/PfDrgFXcGTzRNqR32rwX1xg4MDTxveGi3aNSNHTeUbfx5bsncY1e7HnWtAEOhIKXhEnSkpvplw8joWjgniAQfp9lJ/63j1gyk/OhejO599RntZSuQZ4FuoP8mIso/Mnns8Ie628uYUDROfqUuNhhTjkXkGT63MHrThNfL+LK1dOyuRXCUmBGKEIg2HYc/LCA0PRWK+9fARodJazQuS7C+5aSx3MMoNeeuZb1rZtCckuKxXX73rdL5mYtGdRZkr8gGJxFg75gHo3US2oJ5mF7VBlNRCSLuIRT9eDtaD9XDoo4TCb/toJFa7P1KGUzPLhAF7bSTptK5cUMSh7aDDYja7eg424PxThe4wCCGm+2oP1SDQJxGw5tnYEglkDlrVsQnGY+aiwqv6QiUXSvA2rMjRCjiCMYHmjKEsHM2DEZYspRIm1EINoUBa9ZBX5SNoqUVoKQIRIaBURWEkiHfdKdv3L9u3Qbp8DtvfvkMvv3H32bVfnCSNSvOVpRYhp4iTZkRNRuSMrNVsHljEPQ5SJm9DK/U0tBY0hElVXD4YshKI6FK1iKaXslnTbzybInr5LymM8fY9//2bO6XxmDz2ddTM9W21zzv/HG3IU9ZqONdO0ZPt3g9QZbJmrdMn1qyAEGOQU7FapiSLMiYVglGZYZCaYA+KQvuIV99RkUyk6YXt9h7BvTZpfi13GNdtnX1ypOvvHsu9IVNYvfSsUW5isJYBpuTrMW0vneaJeWMm2LFK9YmTU1OQEURMMqG4Dj/NpISHKJhNQi5DnplEiRTOfK2ryzydR85mVmkQbIxcptOIRKT7c1RoWR79Aslap+7l3BNjOsGWxt0BSn+TZqE8wmeA0JuNxL5y0XXpS5SmxiFKxTH2LAV9jCBMCioCAHJSh7ZhXlIkgkQsioRtw2IhbkCSbJ6iAKBzmb/D3M23XsgmtBOhSOa8KLlc6VrZrC7o8ucqR15QzvygTmgWPYIm5J6ynGhfgVdMo+S2YbJtuZzOGDl0S5jECWU/0vQFKSAHJTdjlIhgc15p1B14wLSEY1DzXtELkIf1s5dMqZNdFU7jlRPkDM33QHAddUAu84ckQVttS+ZZT28muJLFVw4mWb5m0O8MaKeszxmbTyn/s2JfpyjFCApGfQAlmsJ9PsFDBHkR6ERCAJtFIOWMQHvvFCDHy1NR/6COULAbU/EO9vW6fLzC+mII9WS4n2q/fgvAwEu5YHFG+6TPtfFjc1dhnBjbZVadH2LlissCrOOSE8lvk3J+Ft6ak+qv1E9gguU4qOfi1gSv7y7HPsfnIU1hk/u+Xo9gxkKCk0yBt+qtePSqRo6KT95M22t2R4P+mRJOWadgQls41rOrYK2kL0qFx8+eTqyperWulDzhzfrMlNYcGH4pWSMN57BN2qccJKf3NdEQkJjiwvrFxXh+socnL80ATd/JR/fVmTE3asL0XbZiTGSQI09gWUqH8yzlgKecSh1atib+8dciZQty+/ZPXJVAC+feW3hlNsXMJUV2PyjVrdX0E5TE1HqwbdaMAwZ1IIIPyeAJwmQBAESgEOQUHtpAvEpDhsWF+J8zySiElDOMth080IcPNWOSQngCAItw0FsnW+GmzP6pWjszzFzwXEyu2hsy/LZBfvfPT/6f4bYafdQJk342Wyhp84+GApTeQWsOSeTeefI+2gBjQVqCr+/pQCDv1qBht2r8KfNy/DrGyqxc04RytPy0RM04NQlB+6dmQUCgDMsIBInYCM+Thi9BIWDJ+uRlKZlB1u8cjYpky9O52vo7r8//srujfRnmmSi95ys5YNn55cXq4dTytIq6XHnM1qVWhmyBbB3jIJEAwuyVFBrWNT3BjHpdOL1v3dj7ZL1uPveXUhJNsPz4TmoFyzA0FA/xv0v4sRQD5panYhKBICP9f9iP4mNiQClgusuub9jYzBAqmQaizW7pHShe/Thc0nZC8RPAbS1nFtaqrO/oRQNRiZVj7jToUyEougf7MEkTYMA8PuuANDVDgBgBAkPr9+K733nPsgZBn2//R3sx6uRt/N+zNm0CZa0DIzufgiDVi9C4hVNlsopyAigJQb0DLmQMasMag3Dulp7UbKyYlMsHl462n5xJYCOT2mw1JIqKHs7p/nbmnMJkpAnQHNgtfGG9j76tEP4hBQKlCpsLinB93c9jGhnJxzvVSM6PIapoyegmz8PgcEhJGVnoLCwGLXnT6PFG0AhBbz63UoEA3E0OsIo11Aom1kixn0BgaEI0t4+5He0244OhVIPHTp9NvopBr+/Z/94c+3fd9jbTq6RUeGMCKGcma3Gjq7RAIAr91wJwLr0JKxfUIEgsqDRaDDcN4jhBx8GSZEQSBIT//U7JDx+qC/WYOaMMrz1lgYpJIFnd1RAQVM40+EGAAyMOUEp5WRjbXTvtKKM7jFNcU/66qq6rQuXhT7TJGTf0Ttmz1Y9py+YISucprxFrlJS9gT9UfLdWZGFn359Id54/wL0JhMACWk3rYXhgfsQ4QWESAlRXxAZTz8OQ2kxFAo5rL44qrLUsJhVeOZgB1r/J9zeKA+GFsAOntpqLMjiFi5Uvp7oOPa1zzTJ3/Y8LgsPvL84GKLSkpcv/pUgQRkTGGgUDBDmsd7M4of3rERjmxUd4Rj4eOzKJGoWlm/uQEAtBzdug0ohR8rGDSAZOWLhCJiYB0sXpuJ7zzfhnD/+ESsqWoZEnINCq9VoFaGnbGeblNGJ4Lw3X35635Zv7uI/pcFdP3kc2aXKmH/cOoefHMlQZ6SS/ikfXCEJH457cM+yQjAUg2+9dApuQUB6JIilS1bjiaf/gMdefRFOjGAQQbQyPPbsfwcWnQkyXsJo+zG80GhHZ+yTOr6lJANFFjlSMlMIf/8YzZNUi3Jh1T5TwY2de1/c8+kQk6HmhYRv9K20aeYSmubhaW1DyOEO52ZZAADvXhjEt/9wHDaeB00ABrUHbx8+gjhIDPN+LMg0on7Mj0KNiM6AGx6vHy/vexkvjUQwIX2ycJIAFOTkQgyHOD4UhkonQ0ZJyiyzaN0jeZvL/2mIo2Rhfe+pI3sZ3p6uK8gokBRpf3X5px4uzk1h5aKEWv/HrxZsQoBRLcNP9j+H327/LhpeOABGimH17VFIFIvVt0fw0l9fw4nORsw1GUASIiAJGAnEYBMEsLyIGSW56Dj+3nBmXtY+RTSy3tFja/BzsUDBfzzWAWz77Hrw2KHnFEaGX0uCT2SZvAfjUx7m6ROD+GOX9aNvrmMpPPfDNdj8aDUckoA7S8sgKi3QaTSIRwJIxKeQpJOjtFCPipk5gMjD6XTjx8+fxhlvGA/OnoH7VqRh/OC+UHzR179OG7TJAtg/r9q0O/y55RZr69ueW6J+UmlS+uQKFdHV3oO7bijG4V4b3MIVBxZn6mFJZqFXyNAdFdA2MYwNS4w4UluDWr8fIUkCmxCQV0thul4JllWhftiFYYpEGkjcuWomBo++An1OrqJ8jvbVqMsn7+iZcgM4+LnFwqa5M9ICZ07cmPAFU2mNRqY0qBAatUszp+cQRzsnQBAAFYhhxewMfHBxCCOciNFYAtkKCr+4/wYUqXU40zMGgSLhkYC+KI/OQBQ+kgAD4LkdG6B19gpJ6XrSVFRAWmvqFJPN/cOSpeIv+987O/65AF86fbQ/Lo/IWTV3XdwxSUqhmNcaUncXWZjUHD2Lk4Mu2EXAFA6i3eqHXbyilGaHF/myGK5fMg2Ls82obRtHRPr4/JVJwJO3VaFCw2FgPFhn1pImUogx8uxUjqlY8eiiLT89eFXXTk+fnUHQvcE7MElNNPRxYZ7+S/7c6ecotQyrMlV4edNiJBEEft3kwGWe+IQz/7O6F/f9/AAmrU4k0R9PnyWj8cr2jVicLGH04AuQ2z31Xj8O9L7XwDtaR2gx4FkzaXPIrqrkF2Uc19/CPEW8WP8XQgQhzV52ndHvmcjPS5VGglGimA9IB3euJ/50uhVvt48hAOkjt0UJAhc5HnXvd155ywGwtSwLd61bDr6nURz/4AyZXFQiKpbMLRm80JMpHpokXbRTSuzIf7kylxev6lZXP9hJWPc+fAF1bQs0i4uRkpOM1AId5GoGoUgM1uEIF/Yx1XoLURVRqqjL3RM42TiM0WAMkiSBgIRsjRxrZmejojgNLJfAhDXWnJTCFphNpNaglyMWlDDY2IGYxCB4sQ9i5ex9m37z9varAnh4zy8KLBMnOmVmI5NRlOofrBlpVFNhC6mKl18OJqPdwYEkJAASeF4ECBIkSYKmCEiiCAhxcDwgCABBUqDkckBIgJdIXJcnx9xSI0bqhlslubGz6Prsm4e6J9R8f0skULkrZ/09u1yfq0HW1nm7QuZjFHoGQU61J2nbj+/U37SSMs2cA78/jtYhF3xxAVNREVNREXZ/DDZPCKSKgTsuIkgoYQ+LcMQBY6YBx850ISnLhGf+egnxqAiRpJB+2yrJN/+B77hC5CEtK0Kjk6tkPbV3XJVJaAJzFAYT9Col4hK5sOfyxaJg75C1r6bjrEgxAUEC5AoKpiQNtGoGcqUSDi+HYISDSBAoyjODoChotSooaBm23T4HoRAHEAR4iK6BDzvPujvGxslwyyxGTswypyVDk5mNeGBs+lWZJEqpDqjMlg20UkVmWbQrTGn8G0PW/LvNd2xulO17ts81FdYmspIwPGqDnBYRCAuYmgqjf0CC0xvGQP8kaBmJQJRH3SUJkiRCp1YAtAwkLfOIi3avoZj2tQtSuQMq2pzEeaYQ0WgFPr3gXaDp8zUoSRLR9Op3X8otYO9VmEzEwLgHIh9PaHSG4z117Us5QjmV4OKyOJewyMSoWuAFUZWcMU5CFMALgChAEEUqASITiSghJXi3giH8ECVSJonK9NklzRTB31hWkk1ClBANeMWBsdgTczY++chVMUgQhPTaq/se0GJqxCDhkVCE0+h1Stpud9867/pybsIpXoiL8rcnbYFCbWg40945MNlT/NRv1PnlAqG+su1wZwezxvLBL2OOSX1ze+DivMoUMUmv2eZw2ZboTIaqzo5BGI1emIxa35CD/Jk97QfPA09e+yt/Q/UTSzJSqWfiEd/sS00DULAqlM+YBokPwWbzRsyW9Eaf196qYg3dNpvfFgpNRLWsikrPzMzmYpEZsZhQQpPhhV5/nMkvzEVD8yjSUnRQMASUjOJDm4f+wcrNP7/0hdoQredPKTlf3T06JbfTZvcUlJTlE8PjLtAkIAgCFCoKbZeHcNuGJdLImBNO9xQIkScoRgmNRolENAaRoKBUMFCzCpFLoMvplf23fYp6Y+PdjyS+tD7J0T89x2aZfav0WnwjkkjMV8hJkyQKRN/wFLiEhHWrZ2Fg0IZI0A+GphGMJTAtx4ThEaeYmaZ3TTgjdWFO/8KIQ3V227d3XVVL7F9u8v35D79KKcgIV+bkphUE/L45jIzPs6QYyFiMh5jgEAqGBZs9OJhbmNXQ2jw27AoZGrftfNT1/9JM/Gdjy3sSOW8AKJ4JHD4CWO4BHi0jvnBT+x8+c0PISk6iJwAAAABJRU5ErkJggg==);'
    else:
        selected_logo = f"url({convertLogo(arguments['logo'])});"
    page_css = '''
:root {
  --logo: ''' + selected_logo + '''
  --transparent-fingerprint: url(data:image/png;base64,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);
  --fingerprint: url(data:image/png;base64,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);
  --github: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAABhGlDQ1BJQ0MgcHJvZmlsZQAAKJF9kT1Iw0AcxV9TpSIVETuIOASsThbEijhKFYtgobQVWnUwufQLmjQkKS6OgmvBwY/FqoOLs64OroIg+AHi6uKk6CIl/i8ptIjx4Lgf7+497t4BQqPCVLNrElA1y0jFY2I2tyoGXiHAjwGMIioxU0+kFzPwHF/38PH1LsKzvM/9OfqUvMkAn0g8x3TDIt4gntm0dM77xCFWkhTic+IJgy5I/Mh12eU3zkWHBZ4ZMjKpeeIQsVjsYLmDWclQiaeJw4qqUb6QdVnhvMVZrdRY6578hcG8tpLmOs0RxLGEBJIQIaOGMiqwEKFVI8VEivZjHv5hx58kl0yuMhg5FlCFCsnxg//B727NQnTKTQrGgO4X2/4YAwK7QLNu29/Htt08AfzPwJXW9lcbwOwn6fW2Fj4C+reBi+u2Ju8BlzvA0JMuGZIj+WkKhQLwfkbflAMGb4HeNbe31j5OH4AMdbV8AxwcAuNFyl73eHdPZ2//nmn19wOfLXK54rfCmAAAAAlwSFlzAAA3XQAAN10BGYBGXQAAA8hJREFUWMPVmE1oVUcUx3/vkkXc1GxMFsaNtgubB360wY9CqA3YLEpAUHGhVyJXL6KC2grVgJ/oRrC1BMLIqDAtKBJQESRGJHRjlGK1kOjC6MKsEjfqps8u8tycF4ebd+fe93HVDrzFuzN37m9mzvnPOQc+8Zar9kWjVROwEugAPgdagc+k+w0wATwG7gP3/CB8lTmg0SoHrAc2A2uBxpSvFoAh4A9gwA/CYl0BLbDDQL7GUxsFjqcFzaWAWwCcA7rqbF6DwA4/CCeqBjRarQEuA80Z+cAUsMkPwuGKAY1WXcDVCuys2lYA1vlBOJga8APCJULmysDlgb8+IJwN2e4H4aj90IvANQKXysC9Ar4GtgO3gGKVEHeB3TLXeKSvEbgkDDOtITJoX4yM3PaD8AHwANCyy78Ba4BJ4AnwUn5vRbTnAYvFwR4CP9nOYLQaEoG3W16k7NAsQJGTwzErH7P/+EE4arTqBBYBz+L0TPSzFXjpB2HBNae9SUar/pL8NER2L7XdCdR4ijFxOvdfzPNGYdk/Y4Ny7j2Ob7Vl4BRfOvp6SrZYcpK1QJPjhWIGgK45m4RpBrA7Qe3DDAB7gaeO/m4bsMMx8NdqQ6UE+ywAJxxDOgA8o5Un3hh3DAMZivMNx1EvMlrN8USrPAfgs6zo5GTi5veAhV7CHE/8IJzO+Ip76+pMAlwsJvDRWtLHc1aeUfcmN838WgGXZrhBC4C5ruP35KJ32dn3GQL+4Aiap4HnnjjBI8ck24xWczM63p2OIY/8IJwuHfGIY2AzcCqD3fs5IUMcsW0wSYx3Gq0OyqrrsXu7gJMJwwZswD+BF2VsoGg5yyngptHqqxrA8kara0BfQkb5QpjeDzJaHQGOWjdIh1zmp4EtkQn+lgkeAiN+EI7HALUCq4EVwLfAspTFgqN+EB6LBqxngb0S6uSA25I/bJV+G3K5/ApAe4JMnUuQknLR09lZOij3Ym8ksu2TYHWPlCxm2Uk0C4vctRPAhQot4ZAdPUWFuh8YjkCe9oPwNdAJXAH+FfscA86n+OC9CuCGowsqlxe3AP8ALZY9tktWR+luThtEGK02yMKS2iSwxA/CyTSVhTxwx6rJTElOfMPO4IxWuaQKVUrAKaCznLk0xNhOKa0sQTYD14FJo9VzCZHmidfXGtDGwjmDBXnhm4hztACrRDLaqKFCa9UKv3M5mpcQ8Y6LjJwRSalnRveL2PaYa2BDyuTmR6PVRak8rK9h54pS2zng2rWKACNHvtFo9YUk+a2l6yihDQG/S93mYlqw/017B09eQK3tE7jIAAAAAElFTkSuQmCC);
}
.clear { width: 100%; clear: both; }
.card { margin-top: 25px; float: left; width: 400px; height: auto; background-color: ''' + card_colour + '''; border-radius: 5px; padding: 10px; }
.clear-card { position: absolute; }
.card-title { position: relative; justify-items: center; text-align: center; font-family: "Roboto Condensed", sans-serif; font-weight: bold; color: ''' + card_colour + '''; font-size: 40px; margin-bottom: 10px;}
.card-number { position: relative; justify-items: center; text-align: center; font-family: "Roboto Condensed", sans-serif; font-weight: bold; color: ''' + card_colour + '''; }
.grid-container { display: grid; grid-template-columns: auto auto auto; grid-gap: 5px; grid-template-rows: auto; grid-row-gap: 0px; justify-items: center; align-items: center; }
.grid-child { float: left; justify-items: center; align-items: center; padding: 5px; }
.center { display: flex; justify-content: center; position: absolute;  }
.headers > div { float: left; width: 80px; text-align: center; }
.headers > div span { font-size: 30px; color: #fff; font-weight: bold; font-family: Arial, Helvetica, sans-serif;}
.column { float: left; width: 80px; text-align: center; }
.number { padding: 20px 0; border: 2px solid ''' + card_colour + '''; background-color: #fff; font-family: "Roboto Condensed", sans-serif; font-weight: normal; height: 16px; }
.number span { color: #000; font-size: 20px; }
.number span:hover { text-shadow: 0 0 5px rgba(0,0,0,0.5); }
.button {
  background-color: ''' + card_colour + ''';
  border: none;
  color: white;
  padding: 8px 16px;
  border-radius: 8px;
  text-align: center;
  text-decoration: none;
  display: inline-block;
  font-size: 16px;
  margin: 4px 2px;
  cursor: pointer;
  transition-duration: 0.4s;
  width: 100%;
}
.button-clear {
  background-color: white;
  color: black;
  border: 2px solid ''' + card_colour + '''
}
.button-clear:hover {
  background-color: ''' + card_colour + ''';
  color: white;
}
select {
display: block;
margin: 0 auto;
text-align: center;
}
.circle:after { display: flex; content: ''; margin-top: -25px; width: 35px; height: 35px; background: ''' + dauber_colour + '''; border-radius: 50%; position: relative; top: 10%; left: 27%; box-shadow: 0 2px 4px darkslategray, inset 0.1em 0.1em 0.1em 0 rgba(255,255,255,0.5), inset -0.1em -0.1em 0.1em 0 rgba(0,0,0,0.5); -ms-transform: translateY(-50%); transform: translateY(-30%); }
.square:after { display: flex; content: ''; margin-top: -25px; width: 35px; height: 35px; background: ''' + dauber_colour + '''; border-radius: 10%; position: relative; top: 36%; left: 27%; box-shadow: 0 2px 4px darkslategray, inset 0.1em 0.1em 0.1em 0 rgba(255,255,255,0.5), inset -0.1em -0.1em 0.1em 0 rgba(0,0,0,0.5); -ms-transform: translateY(-50%); transform: translateY(-40%); }
.maple-leaf { display: flex; align-items: center; justify-content: center; position: relative; margin-left: 0px; margin-right: 0px; margin-top: 0px; margin-bottom: 0px; }
.maple-leaf:after { position: absolute; content: ''; width: 40px; height: 40px; background: ''' + dauber_colour + '''; -ms-transform: translateY(-20%); transform: translateY(-1%); clip-path: polygon(47% 100%, 48% 70%, 25% 73%, 28% 65%, 7% 47%, 11% 44%, 8% 30%, 20% 32%, 23% 27%, 35% 40%, 32% 13%, 39% 16%, 50% 0, 61% 16%, 68% 13%, 65% 40%, 77% 27%, 80% 32%, 92% 30%, 89% 44%, 93% 47%, 72% 65%, 75% 73%, 52% 70%, 53% 100%); }
.heart:after { display: flex; align-items: center; top: -45px; margin-left: 18%; padding: 0px; position: relative; content: "\\2764"; font-size: 45px; color: ''' + dauber_colour + ''';}
.star:after { display: flex; align-items: center; top: -45px; margin-left: 15%; padding: 0px; position: relative; content: "\\2B50"; font-size: 40px; }
.moon:after { display: flex; align-items: center; top: -45px; margin-left: 13%; padding: 0px; position: relative; content: "\\01F319"; font-size: 40px; }
.unicorn:after { display: flex; align-items: center; top: -45px; margin-left: 17%; padding: 0px; position: relative; content: "\\01F984"; font-size: 40px; }
.clover:after { display: flex; align-items: center; top: -45px; margin-left: 10px; padding: 0px; position: relative; content: "\\01F340"; font-size: 40px;}
.logo:after { display: flex; align-items: center; top: -40px; margin-left: 23%; padding: 0px; position: relative; content: var(--logo); font-size: 40px; }
.x-mark:after { display: flex; content: "\\274C"; margin-top: -20px; margin-left: 3px; position: relative;  font-size: 50px; -ms-transform: translateY(-50%); transform: translateY(-50%); }
.checkmark:after { display: flex; content: "\\2705"; color: green; margin-top: -19px; margin-left: 4px; position: relative; font-size: 48px; font-weight: bold; -ms-transform: translateY(-50%); transform: translateY(-50%);}
.footer { opacity: 0.7; display: flex; justify-content: center; content: var(--github);  margin-left: 49.25%; bottom: 0; height: 40px; width: 40px;}
.footer:hover { opacity: 1.0; }
'''
    close_style = "\n</style>\n"
    close_head = "</head>\n"
    open_body = "<body>\n"
    if arguments['allow_select']:
        select_box = '''
<select id="dauber" name="dauber" class="selectpicker">
  <option selected disabled>Select a dauber</option>
  <option value="checkmark">&#9989; Checkmark</option>
  <option value="circle">&#128308; Circle</option>
  <option value="clover">&#127808; Clover</option>
  <option value="heart">&#10084; Heart</option>
  <option value="maple-leaf">&#127809; Maple Leaf</option>
  <option value="moon">&#127769; Moon</option>
  <option value="square">&#128998; Square</option>
  <option value="star">&#11088; Star</option>
  <option value="unicorn">&#129412; Unicorn</option>
  <option value="x-mark">&#10060; X Marks The Spot</option>
</select>\n
'''
    else:
        select_box = '\n'
    body = '''
<div class="card-title">''' + card_title + '''</div>
<div class="grid-container 1">
<div class="grid-child 1">
<div class="clear"></div>
<div class="card 1">
  <div class="headers">
    <div><span>B</span></div>
    <div><span>I</span></div>
    <div><span>N</span></div>
    <div><span>G</span></div>
    <div><span>O</span></div>
  </div>
  <div class="column 1">
    <div class="number col-1" id="card1-c1"></div>
    <div class="number col-6" id="card1-c6"></div>
    <div class="number col-11" id="card1-c11"></div>
    <div class="number col-16" id="card1-c16"></div>
    <div class="number col-21" id="card1-c21"></div>
  </div>
  <div class="column 2">
    <div class="number col-2" id="card1-c2"></div>
    <div class="number col-7" id="card1-c7"></div>
    <div class="number col-12" id="card1-c12"></div>
    <div class="number col-17" id="card1-c17"></div>
    <div class="number col-22" id="card1-c22"></div>
  </div>
  <div class="column 3">
    <div class="number col-3" id="card1-c3"></div>
    <div class="number col-8" id="card1-c8"></div>
    <div class="number col-13" id="card1-c13"></div>
    <div class="number col-18" id="card1-c18"></div>
    <div class="number col-23" id="card1-c23"></div>
  </div>
  <div class="column 4">
    <div class="number col-4" id="card1-c4"></div>
    <div class="number col-9" id="card1-c9"></div>
    <div class="number col-14" id="card1-c14"></div>
    <div class="number col-19" id="card1-c19"></div>
    <div class="number col-24" id="card1-c24"></div>
  </div>
  <div class="column 5">
    <div class="number col-5" id="card1-c5"></div>
    <div class="number col-10" id="card1-c10"></div>
    <div class="number col-15" id="card1-c15"></div>
    <div class="number col-20" id="card1-c20"></div>
    <div class="number col-25" id="card1-c25"></div>
  </div>
</div>
</div>

<div class="grid-child 2">
<div class="clear"></div>
<div class="card 2">
  <div class="headers">
    <div><span>B</span></div>
    <div><span>I</span></div>
    <div><span>N</span></div>
    <div><span>G</span></div>
    <div><span>O</span></div>
  </div>
  <div class="column 1">
    <div class="number col-1" id="card2-c1"></div>
    <div class="number col-6" id="card2-c6"></div>
    <div class="number col-11" id="card2-c11"></div>
    <div class="number col-16" id="card2-c16"></div>
    <div class="number col-21" id="card2-c21"></div>
  </div>
  <div class="column 2">
    <div class="number col-2" id="card2-c2"></div>
    <div class="number col-7" id="card2-c7"></div>
    <div class="number col-12" id="card2-c12"></div>
    <div class="number col-17" id="card2-c17"></div>
    <div class="number col-22" id="card2-c22"></div>
  </div>
  <div class="column 3">
    <div class="number col-3" id="card2-c3"></div>
    <div class="number col-8" id="card2-c8"></div>
    <div class="number col-13" id="card2-c13"></div>
    <div class="number col-18" id="card2-c18"></div>
    <div class="number col-23" id="card2-c23"></div>
  </div>
  <div class="column 4">
    <div class="number col-4" id="card2-c4"></div>
    <div class="number col-9" id="card2-c9"></div>
    <div class="number col-14" id="card2-c14"></div>
    <div class="number col-19" id="card2-c19"></div>
    <div class="number col-24" id="card2-c24"></div>
  </div>
  <div class="column 5">
    <div class="number col-5" id="card2-c5"></div>
    <div class="number col-10" id="card2-c10"></div>
    <div class="number col-15" id="card2-c15"></div>
    <div class="number col-20" id="card2-c20"></div>
    <div class="number col-25" id="card2-c25"></div>
  </div>
</div>
</div>

<div class="grid-child 3">
<div class="clear"></div>
<div class="card 3">
  <div class="headers">
    <div><span>B</span></div>
    <div><span>I</span></div>
    <div><span>N</span></div>
    <div><span>G</span></div>
    <div><span>O</span></div>
  </div>
  <div class="column 1">
    <div class="number col-1" id="card3-c1"></div>
    <div class="number col-6" id="card3-c6"></div>
    <div class="number col-11" id="card3-c11"></div>
    <div class="number col-16" id="card3-c16"></div>
    <div class="number col-21" id="card3-c21"></div>
  </div>
  <div class="column 2">
    <div class="number col-2" id="card3-c2"></div>
    <div class="number col-7" id="card3-c7"></div>
    <div class="number col-12" id="card3-c12"></div>
    <div class="number col-17" id="card3-c17"></div>
    <div class="number col-22" id="card3-c22"></div>
  </div>
  <div class="column 3">
    <div class="number col-3" id="card3-c3"></div>
    <div class="number col-8" id="card3-c8"></div>
    <div class="number col-13" id="card3-c13"></div>
    <div class="number col-18" id="card3-c18"></div>
    <div class="number col-23" id="card3-c23"></div>
  </div>
  <div class="column 4">
    <div class="number col-4" id="card3-c4"></div>
    <div class="number col-9" id="card3-c9"></div>
    <div class="number col-14" id="card3-c14"></div>
    <div class="number col-19" id="card3-c19"></div>
    <div class="number col-24" id="card3-c24"></div>
  </div>
  <div class="column 5">
    <div class="number col-5" id="card3-c5"></div>
    <div class="number col-10" id="card3-c10"></div>
    <div class="number col-15" id="card3-c15"></div>
    <div class="number col-20" id="card3-c20"></div>
    <div class="number col-25" id="card3-c25"></div>
  </div>
</div>
</div>

</div>

<div class="grid-container 2">

<div class="grid-child 4">
<div class="clear"></div>
<div class="card 4">
  <div class="headers">
    <div><span>B</span></div>
    <div><span>I</span></div>
    <div><span>N</span></div>
    <div><span>G</span></div>
    <div><span>O</span></div>
  </div>
  <div class="column 1">
    <div class="number col-1" id="card4-c1"></div>
    <div class="number col-6" id="card4-c6"></div>
    <div class="number col-11" id="card4-c11"></div>
    <div class="number col-16" id="card4-c16"></div>
    <div class="number col-21" id="card4-c21"></div>
  </div>
  <div class="column 2">
    <div class="number col-2" id="card4-c2"></div>
    <div class="number col-7" id="card4-c7"></div>
    <div class="number col-12" id="card4-c12"></div>
    <div class="number col-17" id="card4-c17"></div>
    <div class="number col-22" id="card4-c22"></div>
  </div>
  <div class="column 3">
    <div class="number col-3" id="card4-c3"></div>
    <div class="number col-8" id="card4-c8"></div>
    <div class="number col-13" id="card4-c13"></div>
    <div class="number col-18" id="card4-c18"></div>
    <div class="number col-23" id="card4-c23"></div>
  </div>
  <div class="column 4">
    <div class="number col-4" id="card4-c4"></div>
    <div class="number col-9" id="card4-c9"></div>
    <div class="number col-14" id="card4-c14"></div>
    <div class="number col-19" id="card4-c19"></div>
    <div class="number col-24" id="card4-c24"></div>
  </div>
  <div class="column 5">
    <div class="number col-5" id="card4-c5"></div>
    <div class="number col-10" id="card4-c10"></div>
    <div class="number col-15" id="card4-c15"></div>
    <div class="number col-20" id="card4-c20"></div>
    <div class="number col-25" id="card4-c25"></div>
  </div>
</div>
</div>

<div class="grid-child 5">
<div class="clear"></div>
<div class="card 5">
  <div class="headers">
    <div><span>B</span></div>
    <div><span>I</span></div>
    <div><span>N</span></div>
    <div><span>G</span></div>
    <div><span>O</span></div>
  </div>
  <div class="column 1">
    <div class="number col-1" id="card5-c1"></div>
    <div class="number col-6" id="card5-c6"></div>
    <div class="number col-11" id="card5-c11"></div>
    <div class="number col-16" id="card5-c16"></div>
    <div class="number col-21" id="card5-c21"></div>
  </div>
  <div class="column 2">
    <div class="number col-2" id="card5-c2"></div>
    <div class="number col-7" id="card5-c7"></div>
    <div class="number col-12" id="card5-c12"></div>
    <div class="number col-17" id="card5-c17"></div>
    <div class="number col-22" id="card5-c22"></div>
  </div>
  <div class="column 3">
    <div class="number col-3" id="card5-c3"></div>
    <div class="number col-8" id="card5-c8"></div>
    <div class="number col-13" id="card5-c13"></div>
    <div class="number col-18" id="card5-c18"></div>
    <div class="number col-23" id="card5-c23"></div>
  </div>
  <div class="column 4">
    <div class="number col-4" id="card5-c4"></div>
    <div class="number col-9" id="card5-c9"></div>
    <div class="number col-14" id="card5-c14"></div>
    <div class="number col-19" id="card5-c19"></div>
    <div class="number col-24" id="card5-c24"></div>
  </div>
  <div class="column 5">
    <div class="number col-5" id="card5-c5"></div>
    <div class="number col-10" id="card5-c10"></div>
    <div class="number col-15" id="card5-c15"></div>
    <div class="number col-20" id="card5-c20"></div>
    <div class="number col-25" id="card5-c25"></div>
  </div>
</div>
</div>

<div class="grid-child 6">
<div class="clear"></div>
<div class="card 6">
  <div class="headers">
    <div><span>B</span></div>
    <div><span>I</span></div>
    <div><span>N</span></div>
    <div><span>G</span></div>
    <div><span>O</span></div>
  </div>
  <div class="column 1">
    <div class="number col-1" id="card6-c1"></div>
    <div class="number col-6" id="card6-c6"></div>
    <div class="number col-11" id="card6-c11"></div>
    <div class="number col-16" id="card6-c16"></div>
    <div class="number col-21" id="card6-c21"></div>
  </div>
  <div class="column 2">
    <div class="number col-2" id="card6-c2"></div>
    <div class="number col-7" id="card6-c7"></div>
    <div class="number col-12" id="card6-c12"></div>
    <div class="number col-17" id="card6-c17"></div>
    <div class="number col-22" id="card6-c22"></div>
  </div>
  <div class="column 3">
    <div class="number col-3" id="card6-c3"></div>
    <div class="number col-8" id="card6-c8"></div>
    <div class="number col-13" id="card6-c13"></div>
    <div class="number col-18" id="card6-c18"></div>
    <div class="number col-23" id="card6-c23"></div>
  </div>
  <div class="column 4">
    <div class="number col-4" id="card6-c4"></div>
    <div class="number col-9" id="card6-c9"></div>
    <div class="number col-14" id="card6-c14"></div>
    <div class="number col-19" id="card6-c19"></div>
    <div class="number col-24" id="card6-c24"></div>
  </div>
  <div class="column 5">
    <div class="number col-5" id="card6-c5"></div>
    <div class="number col-10" id="card6-c10"></div>
    <div class="number col-15" id="card6-c15"></div>
    <div class="number col-20" id="card6-c20"></div>
    <div class="number col-25" id="card6-c25"></div>
  </div>
</div>
</div>

</div>
<a href="https://github.com/digitalsleuth/bingo-card-generator" class="footer"></a>
<script src='https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js'></script>
<script id='rendered-js' >
'''

    js1 = '''
$i = [$card1, $card2, $card3, $card4, $card5, $card6];

$(document).ready(function() {
  for ($card = 1; $card <= 6; $card++) {
    for ($x = 0; $x <= 24; $x++) {
    $("<span>" + $i[($card - 1)][$x] + "</span>").appendTo('#card' + $card + '-c' + ($x + 1));
      }
  }
function toggle(id) {
  var element = document.getElementById(id);
  var dauberChoice = $('select[name=dauber]').val();
  $('select[name=dauber]').prop('disabled', true);
  if (dauberChoice == null)
    { dauberChoice = "''' + dauber_shape + '''" };
  element.classList.toggle(dauberChoice);
};
'''
    js2 = '''
  $('.number').click(function() {
    toggle(this.id);
'''
    js3 = '''
  $('#clear-card').click(function() {
    location.reload();
    });
  });
});
</script>

'''
    if arguments['pdf']:
        print("Generating PDF's, please wait")
    while total <= int(arguments['num']):
        filename = f'{output_path}{str(total)}-{card_colour.strip("#")}.html'
        title = f"<title>CARD {str(total)} </title>\n"
        count = 1
        card_clear = ('<div class="card-number" id="clear-card"><button class="button button-clear">CARD ' +
                      str(total) +
                      ' - CLICK HERE TO CLEAR CARD</button></div>\n')
        free_space = ('$(\".col-13\").html(\'<span style=\"color:' +
                      card_colour +
                      '; font-weight:bold\">FREE</span>\');')
        html = open(filename, 'w')
        html.write(open_head)
        html.write(title)
        html.write(open_style)
        html.write(page_css)
        html.write(close_style)
        html.write(close_head)
        html.write(open_body)
        html.write(card_clear)
        html.write(select_box)
        html.write(body)
        while count <= 6:
            nums = str(genNums())
            html.write('$card' + str(count) + ' = ' + nums + ';\n')
            count += 1
        html.write(js1)
        html.write(free_space)
        html.write(js2)
        html.write(js3)
        html.write("</body></html>")
        html.close()
        if arguments['pdf']:
            pdffile = f'{output_path}{str(total)}-{card_colour.strip("#")}.pdf'
            pdfPrint(filename, pdffile)
        current_count = total
        total += 1
    if current_count == 1:
        print("{} card written".format(str(current_count)))
    elif current_count > 1:
        print("{} cards written".format(str(current_count)))


def convertLogo(logo):
    file_name, file_ext = os.path.splitext(logo)
    basewidth = 40
    img = Image.open(logo)
    wpercent = (basewidth / float(img.size[0]))
    hsize = int((float(img.size[1]) * float(wpercent)))
    img = img.resize((basewidth, hsize), Image.Resampling.LANCZOS)
    resized_logo = f'{file_name}-resized{file_ext}'
    img.save(resized_logo)

    filetype = imghdr.what(resized_logo)
    file_handle = open(resized_logo, 'rb').read()
    b64_logo = base64.b64encode(file_handle).decode('utf-8')
    data_uri = f'data:image/{filetype};base64,{b64_logo}'

    return data_uri

def genNums():
    """Generate random numbers for each column"""
    rand = Random()
    card_array = []
    for _ in range(5):
        b = rand.sample(range(1, 16), 1)[0]
        while b in card_array:
            b = rand.sample(range(1, 16), 1)[0]
        card_array.append(b)
        i = rand.sample(range(16, 31), 1)[0]
        while i in card_array:
            i = rand.sample(range(16, 31), 1)[0]
        card_array.append(i)
        n = rand.sample(range(31, 46), 1)[0]
        while n in card_array:
            n = rand.sample(range(31, 46), 1)[0]
        card_array.append(n)
        g = rand.sample(range(46, 61), 1)[0]
        while g in card_array:
            g = rand.sample(range(46, 61), 1)[0]
        card_array.append(g)
        o = rand.sample(range(61, 76), 1)[0]
        while o in card_array:
            o = rand.sample(range(61, 76), 1)[0]
        card_array.append(o)
    return card_array

def pdfPrint(html_file, out_file):
    """Configure options for printing to PDF"""
    options = {
        'page-size': 'Letter',
        'page-width': '8.5in',
        'page-height': '11in',
        'orientation': 'Landscape',
        'margin-top': '0.5in',
        'margin-right': '0.25in',
        'margin-bottom': '0.5in',
        'margin-left': '0.25in',
        'quiet': ''
    }
    with open(html_file, "r") as html:
        html = html.read().replace(' - CLICK HERE TO CLEAR CARD','')
    html_back = f'{html_file}.html'
    with open(html_back, "w") as backup:
        backup.write(html)
    pdfkit.from_file(html_back, out_file, options=options)
    os.remove(html_back)

def grabNumbers(arguments):
    num = int(arguments['num'])
    print("Extracting numbers from {:d} cards".format(num))
    total = 1
    output_path = arguments['output']
    if arguments['everything'] and not arguments['base_colour']:
        basecolour = arguments['card_colour'].lower()
    else:
        basecolour = arguments['base_colour'].lower()
    pattern = '\$card\d = \[*;*'
    if '.html' not in basecolour:
        basecolour = basecolour + '.html'
    while total <= num:
        input_filename = output_path + os.sep + str(total) + "-" + basecolour
        input_filename = input_filename.replace('#', '')
        input_file = open(input_filename, 'r')
        input_file = input_file.readlines()
        output_filename = output_path + os.sep + str(total) + "-" + basecolour.strip('.html') + '.csv'
        output_file = open(output_filename, 'w+')
        full_sheet = []
        for line in input_file:
            match = re.match(pattern, line)
            if match:
                row_1, row_2, row_3, row_4, row_5, \
                row_6, row_7, row_8, row_9, row_10 = ([] for i in range(10))
                line = re.sub(pattern, '', line)
                line = line.replace("];\n", "").replace(',', '')
                line = line.split()
                for i in line:
                    full_sheet.append(i)

        indices = {
            1: [[0, 1, 2, 3, 4], [25, 26, 27, 28, 29], [50, 51, 52, 53, 54]],
            2: [[5, 6, 7, 8, 9], [30, 31, 32, 33, 34], [55, 56, 57, 58, 59]],
            3: [[10, 11, 12, 13, 14], [35, 36, 37, 38, 39], [60, 61, 62, 63, 64]],
            4: [[15, 16, 17, 18, 19], [40, 41, 42, 43, 44], [65, 66, 67, 68, 69]],
            5: [[20, 21, 22, 23, 24], [45, 46, 47, 48, 49], [70, 71, 72, 73, 74]],
            6: [[75, 76, 77, 78, 79], [100, 101, 102, 103, 104], [125, 126, 127, 128, 129]],
            7: [[80, 81, 82, 83, 84], [105, 106, 107, 108, 109], [130, 131, 132, 133, 134]],
            8: [[85, 86, 87, 88, 89], [110, 111, 112, 113, 114], [135, 136, 137, 138, 139]],
            9: [[90, 91, 92, 93, 94], [115, 116, 117, 118, 119], [140, 141, 142, 143, 144]],
            10: [[95, 96, 97, 98, 99], [120, 121, 122, 123, 124], [145, 146, 147, 148, 149]]
        }

        all_rows = [row_1, row_2, row_3, row_4, row_5, row_6, row_7, row_8, row_9, row_10]
        for idx in range(len(indices)):
            for i in indices.get((idx + 1)):
                (all_rows[idx]).extend(map(full_sheet.__getitem__, i))
            all_rows[idx].insert(5, " ")
            all_rows[idx].insert(11, " ")
        all_rows[2][2] = all_rows[2][8] = all_rows[2][14] = "*"
        all_rows[7][2] = all_rows[7][8] = all_rows[7][14] = "*"
        all_rows.insert(5, [", "])
        for row in all_rows:
            row = ','.join(map(str, row)) + '\n'
            output_file.write(row)
        output_file.close()
        total += 1
    print("Extraction Complete")

def writeToExcel(number_of_csvs, base_filename, excel_name, source_path):
    header = [' ', 'B', 'I', 'N', 'G', 'O',
              ' ', 'B', 'I', 'N', 'G', 'O',
              ' ', 'B', 'I', 'N', 'G', 'O']
    excel_name = source_path + os.sep + excel_name
    call_sheet_header = ['B', 'I', 'N', 'G', 'O']
    call_sheet = NamedStyle(name="call_sheet")
    call_sheet.alignment.horizontal = 'center'
    call_sheet.alignment.vertical = 'center'
    bingo_header = NamedStyle(name="bingo_header")
    bingo_header.font = Font(bold=True, name='Arial', size='15')
    bingo_header.alignment.horizontal = 'center'
    bingo_header.alignment.vertical = 'center'
    bingo_header.fill.start_color = 'FFFFFF'
    bingo_header.fill.end_color = 'FFFFFF'
    bingo_header.fill.fill_type = 'solid'
    called_number = PatternFill(bgColor="FFC000")
    free_space = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    borders = PatternFill(start_color='B2B2B2', end_color='B2B2B2', fill_type='solid')
    alignment = Alignment(horizontal='center', vertical='center')
    writer = Workbook(excel_name)
    call_worksheet = writer.create_sheet('CALL')
    call_worksheet.append(call_sheet_header)
    for csvnum in range(1, number_of_csvs + 1):
        csvfile = (source_path + os.sep + str(csvnum) + '-' + base_filename.strip('.html') + '.csv').lower()
        worksheet = writer.create_sheet(str(csvnum))
        readcsv = open(csvfile, 'r', newline='', encoding='utf-8')
        reader = csv.reader(readcsv)
        for row in reader:
            for item in row:
                try:
                    row[row.index(item)] = int(row[row.index(item)])
                except ValueError:
                    pass
            worksheet.append(row)
        readcsv.close()
        os.remove(csvfile)
    writer.save(excel_name)
    wb = load_workbook(excel_name)
    wb.add_named_style(bingo_header)
    wb.add_named_style(call_sheet)
    ws_call = wb['CALL']
    for row in range(1, 17):
        ws_call.row_dimensions[row].height = 20
    call_font = Font(bold=True, name='Arial', size='20')
    col_B = ws_call.column_dimensions['A']
    col_I = ws_call.column_dimensions['B']
    col_N = ws_call.column_dimensions['C']
    col_G = ws_call.column_dimensions['D']
    col_O = ws_call.column_dimensions['E']
    col_B.font = col_I.font = col_N.font = col_G.font = col_O.font = call_font
    col_B.alignment = col_I.alignment = col_N.alignment = col_G.alignment = col_O.alignment = alignment
    for i, cell in enumerate(ws_call["A1":"E1"]):
        for n, cellObj in enumerate(cell):
            cellObj.style = call_sheet
            cellObj.font = Font(bold=True, name='Arial', size='20', color='FF0000')
    for sheet in range(1, number_of_csvs + 1):
        ws = wb[str(sheet)]
        ws.conditional_formatting.add('B1:B14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(B1,CALL!$A$2:$A$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('C1:C14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(C1,CALL!$B$2:$B$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('D1:D14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(D1,CALL!$C$2:$C$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('E1:E14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(E1,CALL!$D$2:$D$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('F1:F14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(F1,CALL!$E$2:$E$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('H1:H14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(H1,CALL!$A$2:$A$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('I1:I14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(I1,CALL!$B$2:$B$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('J1:J14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(J1,CALL!$C$2:$C$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('K1:K14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(K1,CALL!$D$2:$D$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('L1:L14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(L1,CALL!$E$2:$E$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('N1:N14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(N1,CALL!$A$2:$A$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('O1:O14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(O1,CALL!$B$2:$B$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('P1:P14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(P1,CALL!$C$2:$C$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('Q1:Q14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(Q1,CALL!$D$2:$D$16,1,FALSE)))'], fill=called_number))
        ws.conditional_formatting.add('R1:R14', FormulaRule(formula=['NOT(ISNA(VLOOKUP(R1,CALL!$E$2:$E$16,1,FALSE)))'], fill=called_number))
        for row in range(1, 18):
            ws.row_dimensions[row].height = 20
        column = 1
        while column <= 19:
            col = get_column_letter(column)
            ws.column_dimensions[col].width = 5
            column += 1
        ws.insert_cols(1)
        ws.insert_rows(0)
        ws.insert_rows(1)
        ws.insert_rows(8)
        ws['D5'].fill = ws['J5'].fill = ws['P5'].fill = free_space
        ws['D12'].fill = ws['J12'].fill = ws['P12'].fill = free_space
        ws['D5'].alignment = ws['J5'].alignment = ws['P5'].alignment = alignment
        ws['D12'].alignment = ws['J12'].alignment = ws['P12'].alignment = alignment
        for col, val in enumerate(header, start=1):
            ws.cell(row=2, column=col).value = val
            ws.cell(row=9, column=col).value = val
        for cell in ws["1:1"]:
            cell.fill = borders
        for cell in ws["2:2"]:
            cell.style = bingo_header
        for cell in ws["8:8"]:
            cell.fill = borders
        for cell in ws["9:9"]:
            cell.style = bingo_header
        for cell in ws["15:15"]:
            cell.fill = borders
        for i, cell in enumerate(ws["A1":"A15"]):
            for n, cellObj in enumerate(cell):
                cellObj.fill = borders
        for i, cell in enumerate(ws["G1":"G15"]):
            for n, cellObj in enumerate(cell):
                cellObj.fill = borders
        for i, cell in enumerate(ws["M1":"M15"]):
            for n, cellObj in enumerate(cell):
                cellObj.fill = borders
        for i, cell in enumerate(ws["S1":"S15"]):
            for n, cellObj in enumerate(cell):
                cellObj.fill = borders
        for row in ws.rows:
            for cell in row:
                cell.alignment = alignment
    wb.save(excel_name)

def main():
    """Parse arguments for PDF, card and dauber colour, and dauber shape"""
    arg_parse = argparse.ArgumentParser(
        description='Interactive Bingo Card and PDF Generator v' + str(__version__),
        epilog="If you'd like to see a few other color options, you can visit:\n" + __colour_groups__,
        formatter_class=argparse.RawTextHelpFormatter)
    arg_parse.add_argument('-v', '--version', action='version', version='%(prog)s ' + str(__version__))
    group = arg_parse.add_argument_group(title='positional arguments',
        description='''NUM_OF_CARDS

Customization options:

-e, --everything              Creates HTML files, PDF's and spreadsheet - use -c, -d, and -s for desired customization
                              otherwise your cards will be set to default values.
-p, --pdf                     Convert generated HTML files to PDFs
-o, --output <output_dir>     Choose output directory
-c, --card-colour <colour>    Colour for the card - default is BLUE
-d, --dauber-colour <colour>  Colour for the dauber - default is RED
-s, --dauber-shape <shape>    Shape for the dauber - default is CIRCLE
                              Options are: square, circle, maple-leaf, heart, star, moon, unicorn, clover, logo, checkmark, x-mark
-l, --logo <image_file>       If the logo is chosen for a shape, this is used to point to an image, JPG or PNG to use as a dauber
-a, --allow-select            If chosen, provides a dropdown box on the HTML card for the player to change their dauber at will
''')
    group.add_argument('-p', '--pdf', action='store_true', help=argparse.SUPPRESS)
    group.add_argument('-o', '--output', metavar='', help=argparse.SUPPRESS, required=True)
    group.add_argument('num', metavar='', help=argparse.SUPPRESS, type=int, nargs=1)
    group.add_argument('-c', '--card-colour', help=argparse.SUPPRESS, default='blue')
    group.add_argument('-d', '--dauber-colour', help=argparse.SUPPRESS, default='red')
    group.add_argument('-s', '--dauber-shape', help=argparse.SUPPRESS,
        choices=['square','circle','maple-leaf','heart','star','moon','unicorn','clover','logo', 'checkmark', 'x-mark'], default='circle')
    group.add_argument('-l', '--logo', help=argparse.SUPPRESS)
    group.add_argument('-a', '--allow-select', help=argparse.SUPPRESS, action='store_true')
    group.add_argument('-b', '--base-colour', help=argparse.SUPPRESS)
    group.add_argument('-x', '--excel', help=argparse.SUPPRESS)
    group.add_argument('-e', '--everything', help=argparse.SUPPRESS, action='store_true')

    if len(sys.argv[1:]) == 0:
        arg_parse.print_help()
        arg_parse.exit()

    args = arg_parse.parse_args()
    all_args = vars(args)
    all_args['num'] = all_args['num'][0]
    if all_args['excel'] and all_args['base_colour']:
        grabNumbers(all_args)
        writeToExcel(all_args['num'], all_args['base_colour'], all_args['excel'], all_args['output'])
    elif all_args['excel'] and not all_args['base_colour']:
        print("The Excel option requires the -b, --base-colour value as well")
        raise SystemExit(0)
    elif all_args['everything'] and all_args['excel']:
        all_args['pdf'] = True
        createCard(all_args)
        grabNumbers(all_args)
        writeToExcel(all_args['num'], all_args['card_colour'], all_args['excel'], all_args['output'])
    elif all_args['everything'] and not all_args['excel']:
        all_args['pdf'] = True
        createCard(all_args)
        grabNumbers(all_args)
        writeToExcel(all_args['num'], all_args['card_colour'], str(all_args['card_colour'] + '-cards.xlsx'), all_args['output'])
    else:
        createCard(all_args)


if __name__ == '__main__':
    main()
