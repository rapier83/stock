from kiwoom.kiwoom import *
import sys
from PyQt5.QtWidgets import *


class Main():
    def __init__(self):
        print("Program initiated.")

        self.app = QApplication(sys.argv)
        self.kiwoom = kiwoom()
        self.app.exec_()


if __name__ == '__main__':
    Main()
