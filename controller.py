# -*- coding: utf-8 -*-
"""
Created on Sun Oct 16 05:44:05 2022

@author: tony
"""
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import (QApplication, QMessageBox)
import cv2

from UI2 import Ui_MainWindow
import word_input as word
class MainWindow_controller(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__() # in python3, super(Class, self).xxx = super().xxx
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setup_control()

    def setup_control(self):
        self.ui.pushButton.clicked.connect(self.open_file) 

    def open_file(self):
        filename, filetype = QFileDialog.getOpenFileName(self,
                  "Open file",
                  "./")                 # start path
        txt,nq = word.word_clip(filename)
        print(filename, filetype)
        self.ui.textEdit.setText(filename)
        print(txt)
        
        if txt == "error" :           
            QMessageBox.warning(None, '訊息', '檔案錯誤')
        else:
            text = '工作完成，題數 : ' + str(nq) + "題"
            QMessageBox.information(None, '訊息', text)

    