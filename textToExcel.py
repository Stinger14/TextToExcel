#! D:\Python\python3

from os import listdir
from os.path import isfile, join
import xlwt
import xlrd
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, \
    QInputDialog, QLineEdit, QHBoxLayout, QVBoxLayout, QMessageBox
from PyQt5.QtGui import QIcon, QFont
from subprocess import call

