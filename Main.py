from PyQt5.QtWidgets import QApplication
from core_functions import *
import sys

class Main(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = core_functions()
        self.ui.setupUi(self)
        self.ui.setupOtherAttributes()
        self.show()

def main():

    print("\nNote to self: To run this from commandline, make sure you use the command \"python3.6\", "
          "\nwhich is setup by Pycharm, and whose path is defined in \"~/.bash_profile\"\n")

    # cloudinary.config(
    #     cloud_name="dphxz4d4i",
    #     api_key="521134793219841",
    #     api_secret="OVX7ykjiq37D03ex0zojE4aRqB0"
    # )

    app = QApplication(sys.argv)
    instance = Main()
    sys.exit(app.exec_())
    #app.exec_()
    
if __name__ == '__main__':
    main()
