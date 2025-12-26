import base64
import mimetypes
import os
import re
import subprocess
import sys
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QIcon, QImage, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QCheckBox,
    QComboBox,
    QTabWidget,
    QTextEdit,
)


SUPPORTED_FILTER = "Images (*.png *.jpg *.jpeg *.ico)"
DETAIL_LEVELS = {
    "ต่ำ": 64,
    "กลาง": 128,
    "สูง": 192,
    "สูงมาก (ไฟล์ใหญ่)": 256,
}
INKSCAPE_CANDIDATES = [
    r"C:\Program Files\Inkscape\bin\inkscape.exe",
    r"C:\Program Files\Inkscape\inkscape.exe",
    r"C:\Program Files (x86)\Inkscape\bin\inkscape.exe",
    r"C:\Program Files (x86)\Inkscape\inkscape.exe",
]


def guess_mime_type(path):
    mime, _ = mimetypes.guess_type(path)
    if mime in {"image/png", "image/jpeg", "image/x-icon", "image/vnd.microsoft.icon"}:
        return mime
    lower = path.lower()
    if lower.endswith(".png"):
        return "image/png"
    if lower.endswith(".jpg") or lower.endswith(".jpeg"):
        return "image/jpeg"
    if lower.endswith(".ico"):
        return "image/x-icon"
    return "image/png"


def build_svg(image_bytes, mime_type, width, height):
    b64 = base64.b64encode(image_bytes).decode("ascii")
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        f"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{width}\" height=\"{height}\" "
        f"viewBox=\"0 0 {width} {height}\">\n"
        f"  <image width=\"{width}\" height=\"{height}\" href=\"data:{mime_type};base64,{b64}\" />\n"
        "</svg>\n"
    )


def build_svg_vector(image, width, height, detail_size):
    scaled = image.scaled(
        detail_size,
        detail_size,
        Qt.AspectRatioMode.KeepAspectRatio,
        Qt.TransformationMode.SmoothTransformation,
    )
    w = scaled.width()
    h = scaled.height()
    if w <= 0 or h <= 0:
        raise ValueError("ขนาดรูปภาพไม่ถูกต้อง")

    cell_w = width / w
    cell_h = height / h
    svg_lines = [
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        f"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{width}\" height=\"{height}\" "
        f"viewBox=\"0 0 {width} {height}\">",
    ]

    for y in range(h):
        for x in range(w):
            color = scaled.pixelColor(x, y)
            if color.alpha() == 0:
                continue
            fill = f"#{color.red():02x}{color.green():02x}{color.blue():02x}"
            opacity = color.alpha() / 255
            px = x * cell_w
            py = y * cell_h
            if opacity < 1:
                svg_lines.append(
                    f"  <rect x=\"{px:.3f}\" y=\"{py:.3f}\" width=\"{cell_w:.3f}\" "
                    f"height=\"{cell_h:.3f}\" fill=\"{fill}\" fill-opacity=\"{opacity:.3f}\" />"
                )
            else:
                svg_lines.append(
                    f"  <rect x=\"{px:.3f}\" y=\"{py:.3f}\" width=\"{cell_w:.3f}\" "
                    f"height=\"{cell_h:.3f}\" fill=\"{fill}\" />"
                )

    svg_lines.append("</svg>")
    return "\n".join(svg_lines) + "\n"


def minify_svg_text(text):
    if not text:
        return ""
    compact = re.sub(r"<\?xml[^>]*\?>", "", text)
    compact = re.sub(r">\s+<", "><", compact.strip())
    compact = re.sub(r"\s{2,}", " ", compact)

    def trim_number(match):
        value = float(match.group(0))
        s = f"{value:.3f}"
        return s.rstrip("0").rstrip(".")

    compact = re.sub(r"-?\d+\.\d+", trim_number, compact)
    compact = compact.replace(" />", "/>")
    return compact


def find_inkscape_path():
    for path in INKSCAPE_CANDIDATES:
        if os.path.isfile(path):
            return path
    return ""


class ImageToSvgApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Image to SVG Converter")
        self.resize(920, 620)

        self.image = None
        self.image_path = ""
        self.current_svg_text = ""
        self.inkscape_path = find_inkscape_path()

        self._build_ui()
        self._apply_style()

    def _build_ui(self):
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)

        header = QLabel("แปลงรูปภาพ JPG/PNG เป็นไฟล์ SVG")
        header.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        layout.addWidget(header)

        sub = QLabel("ไฟล์ SVG ที่ได้จะฝังรูปภาพเดิมไว้ภายใน (เหมาะสำหรับนำไปใช้งานต่อ)")
        sub.setFont(QFont("Segoe UI", 10))
        layout.addWidget(sub)

        path_card = QFrame()
        path_card.setObjectName("card")
        path_layout = QGridLayout(path_card)
        path_layout.setHorizontalSpacing(10)
        path_layout.setVerticalSpacing(10)

        self.input_line = QLineEdit()
        self.input_line.setPlaceholderText("เลือกไฟล์รูปภาพ .png .jpg หรือ .ico")
        self.input_line.setReadOnly(True)
        btn_input = QPushButton("เลือกไฟล์")
        btn_input.clicked.connect(self.on_pick_image)

        self.output_line = QLineEdit()
        self.output_line.setPlaceholderText("เลือกไฟล์ผลลัพธ์ .svg")
        self.output_line.setReadOnly(True)
        btn_output = QPushButton("บันทึกเป็น")
        btn_output.clicked.connect(self.on_pick_output)
       

        path_layout.addWidget(QLabel("ไฟล์ต้นฉบับ:"), 0, 0)
        path_layout.addWidget(self.input_line, 0, 1)
        path_layout.addWidget(btn_input, 0, 2)
        path_layout.addWidget(QLabel("ไฟล์ปลายทาง:"), 1, 0)
        path_layout.addWidget(self.output_line, 1, 1)
        path_layout.addWidget(btn_output, 1, 2)

        layout.addWidget(path_card)

        option_card = QFrame()
        option_card.setObjectName("card")
        option_layout = QGridLayout(option_card)
        option_layout.setHorizontalSpacing(10)
        option_layout.setVerticalSpacing(10)

        self.width_line = QLineEdit()
        self.width_line.setPlaceholderText("กว้าง (px)")
        self.height_line = QLineEdit()
        self.height_line.setPlaceholderText("สูง (px)")
        self.keep_ratio = QCheckBox("รักษาอัตราส่วนภาพ")
        self.keep_ratio.setChecked(True)
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Vector (สีเต็ม, พิกเซล)", "Trace (Inkscape)"])
        self.mode_combo.setCurrentText("Trace (Inkscape)")
        self.detail_combo = QComboBox()
        self.detail_combo.addItems(list(DETAIL_LEVELS.keys()))
        self.detail_combo.setCurrentText("สูง")

        option_layout.addWidget(QLabel("ขนาด SVG:"), 0, 0)
        option_layout.addWidget(self.width_line, 0, 1)
        option_layout.addWidget(self.height_line, 0, 2)
        option_layout.addWidget(self.keep_ratio, 1, 1)
        option_layout.addWidget(QLabel("โหมดแปลง:"), 2, 0)
        option_layout.addWidget(self.mode_combo, 2, 1)
        option_layout.addWidget(QLabel("ความละเอียด (เวกเตอร์):"), 2, 2)
        option_layout.addWidget(self.detail_combo, 2, 3)

        self.inkscape_line = QLineEdit()
        self.inkscape_line.setPlaceholderText("เลือก inkscape.exe")
        self.inkscape_line.setText(self.inkscape_path)
        btn_inkscape = QPushButton("เลือก Inkscape")
        btn_inkscape.clicked.connect(self.on_pick_inkscape)
        option_layout.addWidget(QLabel("Inkscape:"), 3, 0)
        option_layout.addWidget(self.inkscape_line, 3, 1, 1, 2)
        option_layout.addWidget(btn_inkscape, 3, 3)

        layout.addWidget(option_card)

        preview_card = QFrame()
        preview_card.setObjectName("card")
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setSpacing(10)

        self.tabs = QTabWidget()
        self.preview = QLabel("พรีวิวรูปภาพ")
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setMinimumHeight(220)

        self.svg_text = QTextEdit()
        self.svg_text.setReadOnly(True)
        self.svg_text.setPlaceholderText("โค้ด SVG จะปรากฏที่นี่หลังแปลงไฟล์")

        self.tabs.addTab(self.preview, "Preview")
        self.tabs.addTab(self.svg_text, "Code")
        preview_layout.addWidget(self.tabs)
        layout.addWidget(preview_card)

        action_row = QHBoxLayout()
        self.btn_convert = QPushButton("แปลงเป็น SVG")
        self.btn_convert.clicked.connect(self.on_convert)
        self._apply_button_icon(self.btn_convert, "BPI.svg")
        self.btn_copy = QPushButton("คัดลอกโค้ด")
        self.btn_copy.clicked.connect(self.on_copy_code)
        self.btn_copy_min = QPushButton("คัดลอกแบบย่อ")
        self.btn_copy_min.clicked.connect(self.on_copy_minified)
        action_row.addWidget(self.btn_convert)
        action_row.addWidget(self.btn_copy)
        action_row.addWidget(self.btn_copy_min)
        action_row.addStretch()
        layout.addLayout(action_row)

        self.setCentralWidget(root)
        self.mode_combo.currentTextChanged.connect(self.on_mode_changed)
        self.on_mode_changed(self.mode_combo.currentText())

    def _apply_style(self):
        self.setStyleSheet(
            """
            QMainWindow {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                            stop:0 #f5f7fb, stop:1 #e7edf5);
            }
            #card {
                background: #ffffff;
                border-radius: 14px;
                border: 1px solid #e2e2e2;
            }
            QLabel {
                color: #1f2a33;
            }
            QLineEdit {
                padding: 8px 10px;
                border: 1px solid #cdd5dc;
                border-radius: 8px;
                background: #fbfdff;
            }
            QPushButton {
                padding: 9px 14px;
                border-radius: 10px;
                border: none;
                background: #226f90;
                color: #ffffff;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #1b5570;
            }
            QCheckBox {
                padding: 4px 6px;
            }
            """
        )

    def _apply_button_icon(self, button, filename):
        icon_path = os.path.join(os.path.dirname(__file__), filename)
        if os.path.isfile(icon_path):
            button.setIcon(QIcon(icon_path))
            button.setIconSize(QSize(20, 20))

    def _message_icon(self, icon_filename):
        icon_path = os.path.join(os.path.dirname(__file__), icon_filename)
        if os.path.isfile(icon_path):
            return QIcon(icon_path)
        return QIcon()

    def _show_message(self, box_type, title, text):
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(text)
        if box_type == "warning":
            box.setIcon(QMessageBox.Icon.Warning)
        elif box_type == "critical":
            box.setIcon(QMessageBox.Icon.Critical)
        else:
            box.setIcon(QMessageBox.Icon.Information)
        icon_file = "Complete.svg" if box_type == "info" else "M.svg"
        icon = self._message_icon(icon_file)
        if not icon.isNull():
            box.setWindowIcon(icon)
        box.exec()

    def on_mode_changed(self, value):
        use_trace = "Trace" in value
        self.detail_combo.setEnabled(not use_trace)

    def on_pick_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "เลือกไฟล์รูปภาพ", "", SUPPORTED_FILTER)
        if not path:
            return

        image = QImage(path)
        if image.isNull():
            QMessageBox.warning(self, "เปิดไฟล์ไม่สำเร็จ", "ไม่สามารถอ่านไฟล์รูปภาพนี้ได้")
            return

        self.image = image
        self.image_path = path
        self.input_line.setText(path)
        self.width_line.setText(str(image.width()))
        self.height_line.setText(str(image.height()))

        base, _ = os.path.splitext(path)
        self.output_line.setText(base + ".svg")
        self._update_preview()

    def on_pick_output(self):
        default_name = "output.svg"
        if self.image_path:
            base = os.path.splitext(self.image_path)[0]
            default_name = base + ".svg"
        path, _ = QFileDialog.getSaveFileName(self, "บันทึกเป็นไฟล์ SVG", default_name, "SVG Files (*.svg)")
        if not path:
            return
        if not path.lower().endswith(".svg"):
            path += ".svg"
        self.output_line.setText(path)

    def on_pick_inkscape(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "เลือกไฟล์ inkscape.exe",
            "",
            "Inkscape (inkscape.exe)",
        )
        if not path:
            return
        self.inkscape_line.setText(path)

    def _update_preview(self):
        if not self.image or self.image.isNull():
            self.preview.setText("พรีวิวรูปภาพ")
            self.preview.setPixmap(QPixmap())
            return
        pixmap = QPixmap.fromImage(self.image)
        target = self.preview.size()
        scaled = pixmap.scaled(target, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.preview.setPixmap(scaled)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_preview()

    def _parse_dimension(self, text, name):
        value = text.strip()
        if not value:
            return None
        if not value.isdigit():
            raise ValueError(f"{name} ต้องเป็นตัวเลขเท่านั้น")
        return int(value)

    def _get_target_size(self):
        if not self.image or self.image.isNull():
            raise ValueError("ยังไม่ได้เลือกรูปภาพ")

        orig_w = self.image.width()
        orig_h = self.image.height()
        if orig_w <= 0 or orig_h <= 0:
            raise ValueError("ขนาดรูปภาพไม่ถูกต้อง")

        width = self._parse_dimension(self.width_line.text(), "ความกว้าง")
        height = self._parse_dimension(self.height_line.text(), "ความสูง")

        if self.keep_ratio.isChecked():
            if width and not height:
                height = max(1, round(width * orig_h / orig_w))
            elif height and not width:
                width = max(1, round(height * orig_w / orig_h))

        if width is None:
            width = orig_w
        if height is None:
            height = orig_h

        return width, height

    def on_convert(self):
        if not self.image_path:
            self._show_message("warning", "ยังไม่ได้เลือกไฟล์", "กรุณาเลือกไฟล์รูปภาพก่อน")
            return

        output_path = self.output_line.text().strip()
        if not output_path:
            self._show_message("warning", "ยังไม่ได้เลือกไฟล์ปลายทาง", "กรุณาระบุไฟล์ SVG ปลายทาง")
            return

        try:
            width, height = self._get_target_size()
        except ValueError as exc:
            self._show_message("warning", "ข้อมูลไม่ถูกต้อง", str(exc))
            return

        use_trace = "Trace" in self.mode_combo.currentText()
        if use_trace:
            inkscape_path = self.inkscape_line.text().strip()
            if not inkscape_path or not os.path.isfile(inkscape_path):
                self._show_message("warning", "ไม่พบ Inkscape", "กรุณาระบุไฟล์ inkscape.exe ให้ถูกต้องก่อน")
                return
            try:
                subprocess.run(
                    [
                        inkscape_path,
                        self.image_path,
                        "--export-type=svg",
                        f"--export-filename={output_path}",
                        "--export-plain-svg",
                        "--actions=select-all;trace-bitmap;export-do",
                    ],
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                )
            except subprocess.CalledProcessError as exc:
                msg = exc.stderr.strip() or exc.stdout.strip() or "ไม่สามารถแปลงด้วย Inkscape ได้"
                self._show_message("critical", "แปลงไม่สำเร็จ", msg)
                return
            try:
                with open(output_path, "r", encoding="utf-8") as fh:
                    svg_text = fh.read()
            except OSError as exc:
                self._show_message("critical", "อ่านไฟล์ไม่สำเร็จ", str(exc))
                return
        else:
            detail_label = self.detail_combo.currentText()
            detail_size = DETAIL_LEVELS.get(detail_label, 64)
            try:
                svg_text = build_svg_vector(self.image, width, height, detail_size)
            except ValueError as exc:
                self._show_message("warning", "ข้อมูลไม่ถูกต้อง", str(exc))
                return
            try:
                with open(output_path, "w", encoding="utf-8") as fh:
                    fh.write(svg_text)
            except OSError as exc:
                self._show_message("critical", "บันทึกไม่สำเร็จ", str(exc))
                return

        if not svg_text:
            self._show_message("critical", "แปลงไม่สำเร็จ", "ไม่พบข้อมูล SVG")
            return

        self.current_svg_text = svg_text
        self.svg_text.setPlainText(svg_text)
        self.tabs.setCurrentWidget(self.svg_text)
        self._show_message("info", "สำเร็จ", f"บันทึกไฟล์ SVG แล้ว\n{output_path}")

    def on_copy_code(self):
        if not self.image_path:
            self._show_message("warning", "ยังไม่ได้เลือกไฟล์", "กรุณาเลือกไฟล์รูปภาพก่อน")
            return
        text = self.svg_text.toPlainText().strip()
        if not text:
            self._show_message("warning", "ยังไม่มีโค้ด", "กรุณาแปลงไฟล์ก่อน แล้วค่อยคัดลอกโค้ด")
            return
        QApplication.clipboard().setText(text)
        self._show_message("info", "คัดลอกแล้ว", "คัดลอกโค้ด SVG ลงคลิปบอร์ดแล้ว")

    def on_copy_minified(self):
        if not self.image_path:
            self._show_message("warning", "ยังไม่ได้เลือกไฟล์", "กรุณาเลือกไฟล์รูปภาพก่อน")
            return
        if not self.current_svg_text:
            self._show_message("warning", "ยังไม่มีโค้ด", "กรุณาแปลงไฟล์ก่อน แล้วค่อยคัดลอกโค้ด")
            return
        compact = minify_svg_text(self.current_svg_text)
        self.svg_text.setPlainText(compact)
        self.tabs.setCurrentWidget(self.svg_text)
        QApplication.clipboard().setText(compact)
        self._show_message("info", "คัดลอกแล้ว", "คัดลอกโค้ด SVG แบบย่อแล้ว")






# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
    # --- ส่วนที่ใช้รันโปรแกรม ---
    #if __name__ == "__main__":
        app = QApplication(sys.argv)
        app.setFont(QFont("Segoe UI", 10))
        window = ImageToSvgApp()
        window.show()
        sys.exit(app.exec())
        
        print(f"--- QUOTA_SAMPLER_INFO: QuotaSamplerApp mainloop finished. ---")

    except Exception as e:
        # ดักจับ Error ที่อาจเกิดขึ้นระหว่างการสร้างหรือรัน App
        print(f"QUOTA_SAMPLER_ERROR: An error occurred during QuotaSamplerApp execution: {e}")
        # แสดง Popup ถ้ามีปัญหา
        if 'root' not in locals() or not root.winfo_exists(): # สร้าง root ชั่วคราวถ้ายังไม่มี
            root_temp = tk.Tk()
            root_temp.withdraw()
            messagebox.showerror("Application Error (Quota Sampler)",
                                f"An unexpected error occurred:\n{e}", parent=root_temp)
            root_temp.destroy()
        else:
            messagebox.showerror("Application Error (Quota Sampler)",
                                f"An unexpected error occurred:\n{e}", parent=root) # ใช้ root ที่มีอยู่ถ้าเป็นไปได้
        sys.exit(f"Error running QuotaSamplerApp: {e}") # อาจจะ exit หรือไม่ก็ได้ ขึ้นกับการออกแบบ


# --- ส่วน Run Application เมื่อรันไฟล์นี้โดยตรง (สำหรับ Test) ---
if __name__ == "__main__":
    print("--- Running QuotaSamplerApp.py directly for testing ---")
    # (ถ้ามีการตั้งค่า DPI ด้านบน มันจะทำงานอัตโนมัติ)

    # เรียกฟังก์ชัน Entry Point ที่เราสร้างขึ้น
    run_this_app()

    print("--- Finished direct execution of QuotaSamplerApp.py ---")
# <<< END OF CHANGES >>>