import sys
import Paketmanager

if __name__ == '__main__':
    app = Paketmanager.QApplication([])
    application = Paketmanager.GUI()
    sys.exit(app.exec_())
