import sys
import traceback

try:
    from PyQt5.QtWidgets import QApplication
    from main import MainWindow
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
except Exception as e:
    print("=" * 80)
    print("ERROR:", str(e))
    print("=" * 80)
    traceback.print_exc()
    print("=" * 80)
    input("Press Enter to exit...")
