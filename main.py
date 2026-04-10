import sys

from PySide6.QtWidgets import QApplication

from app.core.config import init_database
from app.ui.main_window import MainWindow


def main() -> None:
    init_database()

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()