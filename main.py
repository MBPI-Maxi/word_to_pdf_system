from PyQt6.QtWidgets import QApplication
from ConverterWorker import ConverterWorker
from WordToPdfConverter import WordToPdfConverter

import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WordToPdfConverter(converter_worker=ConverterWorker)
    window.show()
    sys.exit(app.exec())