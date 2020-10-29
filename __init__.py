import sys
sys.path.append('C:\\Users\\rapie\\OneDrive\\repos\\source\\python\\stock')

from kiwoom.kiwoom import *
from PyQt5.QtWidgets import *


class Main():
    def __init__(self):
        print("Program initiated.")

        self.app = QApplication(sys.argv)
        self.kiwoom = Kiwoom()
        self.app.exec_()


if __name__ == '__main__':
    Main()
