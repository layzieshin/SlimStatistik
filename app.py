import sys, os
from PyQt6.QtWidgets import QApplication
from controller.main_controller import MainController
from ui.main_window import MainWindow
from util.paths import resource_path
import sys, traceback, os, datetime


def main():
    # Konfig extern neben der EXE anlegen/verwenden
    settings_path = os.path.join(os.getcwd(), "config", "settings.ini")
    os.makedirs(os.path.dirname(settings_path), exist_ok=True)

    # mapping.json kommt aus den App-Ressourcen (funktioniert auch in OneFile)
    mapping_path = resource_path("config/mapping.json")

    ctrl = MainController(settings_path, mapping_path)

    app = QApplication(sys.argv)
    w = MainWindow(ctrl)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
