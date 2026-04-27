from __future__ import annotations

import os
import sys
from pathlib import Path

from PyQt6.QtCore import QObject, QSize, QThread, Qt, pyqtSignal
from PyQt6.QtGui import QFont, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QProgressBar,
    QVBoxLayout,
    QWidget,
)

from .convert_txt_sps_to_sav import convert


class ConvertWorker(QObject):
    log = pyqtSignal(str)
    finished = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, txt_path: Path, sps_paths: list[Path], output_path: Path) -> None:
        super().__init__()
        self.txt_path = txt_path
        self.sps_paths = sps_paths
        self.output_path = output_path

    def run(self) -> None:
        try:
            result = convert(
                self.txt_path,
                self.sps_paths,
                self.output_path,
                logger=self.log.emit,
                row_compress=True,
                compress=False,
            )
            self.finished.emit(result)
        except Exception as exc:
            self.failed.emit(str(exc))


class SpssConverterWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.thread: QThread | None = None
        self.worker: ConvertWorker | None = None

        self.setWindowTitle("SPSS For Lychee Builder")
        try:
            icon_path = Path(__file__).with_name("logo.svg")
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
            else:
                print(f"WARNING: Logo file not found at {icon_path}")
        except Exception as e:
            print(f"WARNING: Could not load window icon: {e}")
        self.resize(1040, 760)
        self.setMinimumSize(900, 680)

        self.txt_input = QLineEdit()
        self.txt_input.setPlaceholderText("Select text data file")
        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("Choose where to save the .sav file")
        self.sps_list = QListWidget()
        self.log_box = QPlainTextEdit()
        self.progress = QProgressBar()
        self.convert_button = QPushButton("Convert to SAV")
        self.clear_sps_button = QPushButton("Clear")
        self.remove_sps_button = QPushButton("Remove")
        self.add_sps_button = QPushButton("Add SPS Files")
        self.txt_button = QPushButton("Browse Text Data")
        self.output_button = QPushButton("Save As")

        self.build_ui()
        self.apply_style()

    def build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(26, 22, 26, 22)
        layout.setSpacing(16)

        logo = QLabel()
        logo.setObjectName("logo")
        logo_path = Path(__file__).with_name("logo.svg")
        try:
            if logo_path.exists():
                logo.setPixmap(QPixmap(str(logo_path)).scaled(58, 58, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            else:
                print(f"WARNING: UI Logo file not found at {logo_path}")
                logo.setText("📊")
                logo.setStyleSheet("font-size: 40px;")
        except Exception as e:
            print(f"WARNING: Could not load UI logo: {e}")
            logo.setText("📊")
            logo.setStyleSheet("font-size: 40px;")

        title = QLabel("SPSS For Lychee Builder")
        title.setObjectName("title")
        subtitle = QLabel("Build IBM SPSS-ready .sav files for Lychee syntax exports.")
        subtitle.setObjectName("subtitle")

        title_stack = QVBoxLayout()
        title_stack.setSpacing(2)
        title_stack.addWidget(title)
        title_stack.addWidget(subtitle)

        header_frame = QFrame()
        header_frame.setObjectName("header")
        header = QHBoxLayout(header_frame)
        header.setContentsMargins(18, 16, 18, 16)
        header.setSpacing(14)
        header.addWidget(logo)
        header.addLayout(title_stack)
        header.addStretch(1)
        layout.addWidget(header_frame)

        panel = QFrame()
        panel.setObjectName("panel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(16, 16, 16, 16)
        panel_layout.setSpacing(12)

        self.add_sps_button.setIcon(QIcon.fromTheme("list-add"))
        self.add_sps_button.setIconSize(QSize(18, 18))
        self.add_sps_button.clicked.connect(self.pick_sps)
        self.remove_sps_button.clicked.connect(self.remove_selected_sps)
        self.clear_sps_button.clicked.connect(self.sps_list.clear)

        sps_card = QFrame()
        sps_card.setObjectName("pathCard")
        sps_card_layout = QVBoxLayout(sps_card)
        sps_card_layout.setContentsMargins(16, 14, 16, 16)
        sps_card_layout.setSpacing(12)

        sps_header = QHBoxLayout()
        sps_label = QLabel("1. SPSS syntax files")
        sps_label.setObjectName("cardTitle")
        sps_hint = QLabel("Add one or more .sps files")
        sps_hint.setObjectName("hint")
        sps_header.addWidget(sps_label)
        sps_header.addWidget(sps_hint)
        sps_header.addStretch(1)
        sps_header.addWidget(self.add_sps_button)
        sps_header.addWidget(self.remove_sps_button)
        sps_header.addWidget(self.clear_sps_button)
        sps_card_layout.addLayout(sps_header)
        sps_card_layout.addWidget(self.sps_list)
        panel_layout.addWidget(sps_card)

        self.txt_button.clicked.connect(self.pick_txt)
        txt_card = QFrame()
        txt_card.setObjectName("pathCard")
        txt_card_layout = QVBoxLayout(txt_card)
        txt_card_layout.setContentsMargins(16, 14, 16, 16)
        txt_card_layout.setSpacing(10)
        txt_label = QLabel("2. Text data")
        txt_label.setObjectName("cardTitle")
        txt_row = QHBoxLayout()
        txt_row.setSpacing(12)
        txt_row.addWidget(self.txt_input, 1)
        txt_row.addWidget(self.txt_button)
        txt_card_layout.addWidget(txt_label)
        txt_card_layout.addLayout(txt_row)
        panel_layout.addWidget(txt_card)

        self.output_button.clicked.connect(self.pick_output)
        output_card = QFrame()
        output_card.setObjectName("pathCard")
        output_card_layout = QVBoxLayout(output_card)
        output_card_layout.setContentsMargins(16, 14, 16, 16)
        output_card_layout.setSpacing(10)
        output_label = QLabel("3. Output SAV")
        output_label.setObjectName("cardTitle")
        output_row = QHBoxLayout()
        output_row.setSpacing(12)
        output_row.addWidget(self.output_input, 1)
        output_row.addWidget(self.output_button)
        output_card_layout.addWidget(output_label)
        output_card_layout.addLayout(output_row)
        panel_layout.addWidget(output_card)

        layout.addWidget(panel)

        log_header = QHBoxLayout()
        status_label = QLabel("Run log")
        status_label.setObjectName("sectionTitle")
        log_header.addWidget(status_label)
        log_header.addStretch(1)
        layout.addLayout(log_header)

        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(150)
        layout.addWidget(self.log_box)

        footer = QHBoxLayout()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        footer.addWidget(self.progress, 1)
        self.convert_button.clicked.connect(self.start_convert)
        footer.addWidget(self.convert_button)
        layout.addLayout(footer)

    def apply_style(self) -> None:
        QApplication.instance().setFont(QFont("Segoe UI", 10))
        self.setStyleSheet(
            """
            QMainWindow {
                background: #0b1020;
            }
            QLabel {
                color: #e5edf8;
            }
            QLabel#title {
                font-size: 30px;
                font-weight: 700;
                color: #f8fafc;
            }
            QLabel#subtitle {
                color: #9fb0c8;
                font-size: 13px;
            }
            QLabel#sectionTitle {
                font-weight: 700;
                color: #f8fafc;
                font-size: 13px;
            }
            QLabel#cardTitle {
                color: #f8fafc;
                font-weight: 700;
                font-size: 13px;
            }
            QLabel#hint {
                color: #8ea3bf;
                font-size: 12px;
            }
            QFrame#header {
                background: #121a2c;
                border: 1px solid #22304a;
                border-radius: 8px;
            }
            QFrame#panel {
                background: #0f172a;
                border: 1px solid #27364f;
                border-radius: 8px;
            }
            QFrame#pathCard {
                background: #111827;
                border: 1px solid #2a3a55;
                border-radius: 8px;
            }
            QLineEdit, QListWidget, QPlainTextEdit {
                background: #0a1222;
                color: #e5edf8;
                border: 1px solid #3b4b66;
                border-radius: 6px;
                padding: 10px;
                selection-background-color: #2563eb;
            }
            QLineEdit:focus, QListWidget:focus, QPlainTextEdit:focus {
                border-color: #38bdf8;
            }
            QListWidget {
                min-height: 118px;
                alternate-background-color: #111c31;
            }
            QListWidget::item {
                padding: 7px;
                color: #dce7f7;
            }
            QListWidget::item:selected {
                background: #1d4ed8;
                color: #ffffff;
            }
            QPlainTextEdit {
                color: #d7e5f8;
                background: #070c18;
                font-family: Consolas, "Cascadia Mono", monospace;
                font-size: 11px;
            }
            QPushButton {
                background: #172033;
                color: #f8fafc;
                border: 1px solid #334155;
                border-radius: 6px;
                padding: 9px 15px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #21304b;
                border-color: #60a5fa;
            }
            QPushButton:pressed {
                background: #1e3a8a;
            }
            QPushButton:disabled {
                color: #64748b;
                background: #111827;
                border-color: #263247;
            }
            QPushButton#primary {
                background: #2563eb;
                color: white;
                border-color: #60a5fa;
                min-width: 136px;
            }
            QPushButton#primary:hover {
                background: #1d4ed8;
            }
            QProgressBar {
                height: 16px;
                border: 1px solid #334155;
                border-radius: 6px;
                background: #111827;
                text-align: center;
            }
            QProgressBar::chunk {
                border-radius: 5px;
                background: #38bdf8;
            }
            """
        )
        self.convert_button.setObjectName("primary")

    def add_sps_item(self, path: Path) -> None:
        resolved = str(path.resolve())
        for index in range(self.sps_list.count()):
            if self.sps_list.item(index).data(Qt.ItemDataRole.UserRole) == resolved:
                return
        item = QListWidgetItem(path.name)
        item.setToolTip(resolved)
        item.setData(Qt.ItemDataRole.UserRole, resolved)
        self.sps_list.addItem(item)

    def pick_txt(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select text data",
            str(Path.cwd()),
            "Text data (*.txt *.tsv *.dat);;All files (*.*)",
        )
        if path:
            self.txt_input.setText(path)

    def pick_sps(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select SPSS syntax files",
            str(Path.cwd()),
            "SPSS syntax (*.sps);;All files (*.*)",
        )
        for path in paths:
            self.add_sps_item(Path(path))

    def pick_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save SPSS file",
            self.output_input.text() or str(Path.cwd() / "SPSS.sav"),
            "SPSS data (*.sav)",
        )
        if path:
            output_path = Path(path)
            if output_path.suffix.lower() != ".sav":
                output_path = output_path.with_suffix(".sav")
            self.output_input.setText(str(output_path))

    def remove_selected_sps(self) -> None:
        for item in self.sps_list.selectedItems():
            self.sps_list.takeItem(self.sps_list.row(item))

    def validate_inputs(self) -> tuple[Path, list[Path], Path] | None:
        txt_path = Path(self.txt_input.text().strip())
        output_path = Path(self.output_input.text().strip())
        sps_paths = [
            Path(self.sps_list.item(index).data(Qt.ItemDataRole.UserRole))
            for index in range(self.sps_list.count())
        ]

        if not txt_path.is_file():
            QMessageBox.warning(self, "Missing text data", "Please select a valid text data file.")
            return None
        if not sps_paths:
            QMessageBox.warning(self, "Missing syntax", "Please add at least one .sps syntax file.")
            return None
        missing_sps = [str(path) for path in sps_paths if not path.is_file()]
        if missing_sps:
            QMessageBox.warning(self, "Missing syntax file", "\n".join(missing_sps[:5]))
            return None
        if not output_path.parent.exists():
            QMessageBox.warning(self, "Output folder not found", str(output_path.parent))
            return None
        if output_path.suffix.lower() != ".sav":
            output_path = output_path.with_suffix(".sav")
            self.output_input.setText(str(output_path))

        return txt_path, sps_paths, output_path

    def start_convert(self) -> None:
        inputs = self.validate_inputs()
        if inputs is None:
            return

        txt_path, sps_paths, output_path = inputs
        self.set_running(True)
        self.log_box.clear()
        self.log("Starting conversion...")
        self.log(f"Text data: {txt_path}")
        self.log(f"SPS files: {len(sps_paths)}")
        self.log(f"Output: {output_path}")
        self.log("Compression: Compatible")

        self.thread = QThread()
        self.worker = ConvertWorker(txt_path, sps_paths, output_path)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.log.connect(self.log)
        self.worker.finished.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def on_finished(self, result: dict) -> None:
        self.set_running(False)
        self.log("")
        self.log(f"Done: {result['output']}")
        self.log(f"Rows: {result['rows']:,}")
        self.log(f"Columns: {result['columns']:,}")
        self.log(f"Variable labels: {result['variable_labels']:,}")
        self.log(f"Value label sets: {result['value_label_sets']:,}")
        self.log(f"MRSETS added: {result['mrsets']:,}")
        self.log(f"File size: {result['file_size_kb']:,} KB")
        QMessageBox.information(self, "Conversion complete", f"Saved:\n{result['output']}")

    def on_failed(self, message: str) -> None:
        self.set_running(False)
        self.log("")
        self.log(f"Error: {message}")
        QMessageBox.critical(self, "Conversion failed", message)

    def set_running(self, running: bool) -> None:
        self.convert_button.setDisabled(running)
        self.progress.setRange(0, 0 if running else 1)
        self.progress.setValue(0 if running else 1)

    def log(self, message: str) -> None:
        self.log_box.appendPlainText(message)


def main() -> int:
    app = QApplication(sys.argv)
    window = SpssConverterWindow()
    window.show()
    return app.exec()




# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน SPSS Converter GUI.
    """
    print(f"--- SPSS_CONVERTER_INFO: Starting 'SPSS Converter GUI' via run_this_app() ---")
    print(f"--- SPSS_CONVERTER_INFO: Working dir: {working_dir} ---")
    print(f"--- SPSS_CONVERTER_INFO: Current dir: {os.getcwd()} ---")
    try:
        # --- ส่วนที่ใช้รันโปรแกรม ---
        print(f"--- SPSS_CONVERTER_INFO: Creating QApplication ---")
        app = QApplication(sys.argv)
        app.setFont(QFont("Segoe UI", 10))
        print(f"--- SPSS_CONVERTER_INFO: Creating SpssConverterWindow ---")
        window = SpssConverterWindow()
        print(f"--- SPSS_CONVERTER_INFO: Showing window ---")
        window.show()
        print(f"--- SPSS_CONVERTER_INFO: Starting app.exec() ---")
        sys.exit(app.exec())

        print(f"--- SPSS_CONVERTER_INFO: Application finished successfully ---")

    except Exception as e:
        # ดักจับ Error ที่อาจเกิดขึ้นระหว่างการสร้างหรือรัน App
        print(f"SPSS_CONVERTER_ERROR: An error occurred during SPSS Converter execution: {e}")
        import traceback
        traceback.print_exc()
        QMessageBox.critical(
            None, "Application Error",
            f"An unexpected error occurred:\n{e}")
        sys.exit(1)


# --- ส่วน Run Application เมื่อรันไฟล์นี้โดยตรง (สำหรับ Test) ---
if __name__ == "__main__":
    print("--- Running SPSS Converter GUI directly for testing ---")
    # เรียกฟังก์ชัน Entry Point ที่เราสร้างขึ้น
    run_this_app()
    print("--- Finished direct execution of SPSS Converter GUI ---")
# <<< END OF CHANGES >>>
