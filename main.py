import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QLabel, QPushButton, QProgressBar,
    QFileDialog, QListWidget, QDialog, QLineEdit
)
from PyQt6.QtCore import Qt
from worker import Worker


# ---------- Year Input Dialog ----------
class YearInputDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Enter Year Values")

        self.inputs = {}
        layout = QVBoxLayout()

        self.years = [2017, 2018, 2019, 2020, 2021, 2022]

        for year in self.years:
            label = QLabel(f"{year} (x,y):")
            inp = QLineEdit()
            inp.setPlaceholderText("e.g. 10,10")

            layout.addWidget(label)
            layout.addWidget(inp)

            self.inputs[year] = inp

        btn = QPushButton("OK")
        btn.clicked.connect(self.accept)

        layout.addWidget(btn)
        self.setLayout(layout)

    def get_values(self):
        year_map = {}

        for year, widget in self.inputs.items():
            text = widget.text().strip()

            if text:
                try:
                    x, y = map(int, text.split(","))
                    year_map[year] = (x, y)
                except:
                    pass

        return year_map


# ---------- Main App ----------
class App(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Processor ⚡")
        self.resize(450, 400)

        self.files = []

        self.status = QLabel("No files selected")
        self.progress = QProgressBar()
        self.file_list = QListWidget()

        self.select_btn = QPushButton("Select Files")
        self.process_btn = QPushButton("Process")

        layout = QVBoxLayout()
        layout.addWidget(self.file_list)
        layout.addWidget(self.status)
        layout.addWidget(self.progress)
        layout.addWidget(self.select_btn)
        layout.addWidget(self.process_btn)

        self.setLayout(layout)

        self.select_btn.clicked.connect(self.select_files)
        self.process_btn.clicked.connect(self.run)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel Files", "", "Excel Files (*.xlsx *.xls)"
        )

        if files:
            self.files = files
            self.file_list.clear()

            for f in files:
                self.file_list.addItem(os.path.basename(f))

            self.status.setText(f"{len(files)} files selected")

    def run(self):
        if not self.files:
            self.status.setText("❌ No files")
            return

        dialog = YearInputDialog()

        if dialog.exec():
            year_map = dialog.get_values()

            if not year_map:
                self.status.setText("❌ Enter valid year values")
                return

            self.status.setText("⏳ Processing...")
            self.process_btn.setEnabled(False)

            self.worker = Worker(self.files, year_map)
            self.worker.progress.connect(self.progress.setValue)
            self.worker.done.connect(self.done)
            self.worker.start()

    def done(self):
        self.status.setText("✅ Done (Check project folder)")
        self.process_btn.setEnabled(True)


# ---------- Run ----------
app = QApplication(sys.argv)
window = App()
window.show()
sys.exit(app.exec())