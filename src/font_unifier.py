import sys
import os

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QFrame, QComboBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from docx import Document
from docx.oxml.ns import qn as docx_qn
from openpyxl import load_workbook
from pptx import Presentation
from pptx.oxml.ns import qn as pptx_qn
from lxml import etree


# --- Font helpers (handle East Asian / complex scripts) ---

def _set_docx_run_font(run, font_name):
    """Word: set latin + eastAsia + complex-script fonts on a run."""
    rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    rfonts.set(docx_qn('w:ascii'), font_name)
    rfonts.set(docx_qn('w:hAnsi'), font_name)
    rfonts.set(docx_qn('w:eastAsia'), font_name)
    rfonts.set(docx_qn('w:cs'), font_name)


def _set_pptx_run_font(run, font_name):
    """PowerPoint: set latin + eastAsian + complex-script typefaces on a run."""
    run.font.name = font_name
    rPr = run.font._rPr
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        elem = rPr.find(pptx_qn(tag))
        if elem is None:
            elem = etree.SubElement(rPr, pptx_qn(tag))
        elem.set('typeface', font_name)


def _set_pptx_text_frame_fonts(text_frame, font_name):
    """Apply the pptx font helper to every run in a text frame."""
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            _set_pptx_run_font(run, font_name)


# --- Core Logic for Font Changing ---

def _set_docx_font(paragraphs, font_name):
    """Apply the docx font helper to every run across the given paragraphs."""
    for para in paragraphs:
        for run in para.runs:
            _set_docx_run_font(run, font_name)


def change_word_font(path, new_font_name):
    """Changes the font for all text in a .docx file."""
    doc = Document(path)
    _set_docx_font(doc.paragraphs, new_font_name)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _set_docx_font(cell.paragraphs, new_font_name)
    return doc


def change_excel_font(path, new_font_name):
    """Changes the font for all cells in a .xlsx file, preserving other style."""
    workbook = load_workbook(path)
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = cell.font.copy(name=new_font_name)
    return workbook


def _process_chart_fonts(chart, font_name):
    """Set fonts on a chart's title and axis titles.

    Wrapped defensively: a malformed chart must never crash the whole file.
    python-pptx exposes category_axis/value_axis (not x_axis/y_axis). Per-point
    data-label fonts are not reliably settable, so they are skipped.
    """
    try:
        if chart.has_title:
            _set_pptx_text_frame_fonts(chart.chart_title.text_frame, font_name)
        for axis_attr in ('category_axis', 'value_axis', 'series_axis'):
            axis = getattr(chart, axis_attr, None)
            if axis is not None and getattr(axis, 'has_title', False):
                try:
                    _set_pptx_text_frame_fonts(
                        axis.axis_title.text_frame, font_name)
                except Exception:
                    pass
    except Exception:
        pass


def change_ppt_font(path, new_font_name):
    """Changes the font for all text in a .pptx file.

    Covers text frames, tables, charts (title/axes) and nested groups.
    """

    def process_shape_text(shape):
        """Recursively processes text in a shape, including nested groups."""
        if getattr(shape, 'has_text_frame', False) and shape.has_text_frame:
            _set_pptx_text_frame_fonts(shape.text_frame, new_font_name)
        if getattr(shape, 'has_table', False) and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    _set_pptx_text_frame_fonts(cell.text_frame, new_font_name)
        if getattr(shape, 'has_chart', False) and shape.has_chart:
            _process_chart_fonts(shape.chart, new_font_name)
        if getattr(shape, 'has_group', False) and shape.has_group:
            for sub_shape in shape.shapes:
                process_shape_text(sub_shape)

    prs = Presentation(path)
    for slide in prs.slides:
        for shape in slide.shapes:
            process_shape_text(shape)
    return prs


# Extension -> handler (case-insensitive dispatch in process_office_file)
_FONT_CHANGERS = {
    ".docx": change_word_font,
    ".xlsx": change_excel_font,
    ".pptx": change_ppt_font,
}


def process_office_file(path, font_name):
    """Process a single Office file and save the modified copy.

    Returns the output path. Raises ValueError on unsupported extensions.
    Case-insensitive on the extension.
    """
    file_dir, file_name = os.path.split(path)
    name, ext = os.path.splitext(file_name)
    output_path = os.path.join(file_dir, f"{name}_modified{ext}")

    changer = _FONT_CHANGERS.get(ext.lower())
    if changer is None:
        raise ValueError(f"Unsupported file type: {ext}")
    changer(path, font_name).save(output_path)
    return output_path


# --- Background worker (keeps the GUI responsive on large files) ---

class FontProcessingWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, path, font_name):
        super().__init__()
        self._path = path
        self._font_name = font_name

    def run(self):
        try:
            output_path = process_office_file(self._path, self._font_name)
            self.finished.emit(output_path)
        except Exception as e:
            self.error.emit(str(e))


# --- GUI Application ---

class FontUnifierApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Font Unifier")
        self.setGeometry(100, 100, 500, 250)

        self.file_path = ""
        self.font_name = "Meiryo UI"
        self._worker = None

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        layout = QVBoxLayout(central_widget)

        # File Selection Frame
        file_frame = QFrame()
        file_layout = QHBoxLayout(file_frame)
        file_layout.setContentsMargins(10, 10, 10, 10)

        file_label = QLabel("File:")
        file_label.setFixedWidth(60)
        file_layout.addWidget(file_label)

        self.file_entry = QLineEdit()
        self.file_entry.setReadOnly(True)
        file_layout.addWidget(self.file_entry)

        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_button)

        layout.addWidget(file_frame)

        # Font Selection Frame
        font_frame = QFrame()
        font_layout = QHBoxLayout(font_frame)
        font_layout.setContentsMargins(10, 10, 10, 10)

        font_label = QLabel("Target Font:")
        font_label.setFixedWidth(80)
        font_layout.addWidget(font_label)

        self.font_entry = QComboBox()
        self.font_entry.addItems([
            "Arial", "Calibri", "Times New Roman", "Verdana", "Tahoma",
            "Georgia", "Comic Sans MS", "Impact", "Courier New",
            "Lucida Sans Unicode", "Meiryo UI", "MS Gothic", "MS Mincho",
            "Meiryo", "Yu Gothic", "Yu Mincho"
        ])
        self.font_entry.setCurrentText(self.font_name)
        font_layout.addWidget(self.font_entry)

        layout.addWidget(font_frame)

        # Action Frame
        action_frame = QFrame()
        action_layout = QVBoxLayout(action_frame)
        action_layout.setContentsMargins(10, 20, 10, 20)

        self.start_button = QPushButton("Start Processing")
        self.start_button.setFixedSize(150, 40)
        self.start_button.clicked.connect(self.process_file)
        action_layout.addWidget(self.start_button,
                                alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(action_frame)

        # Status Label
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: green;")
        layout.addWidget(self.status_label,
                         alignment=Qt.AlignmentFlag.AlignCenter)

    def _set_status(self, text, color):
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"color: {color};")

    def browse_file(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilters([
            "Office Files (*.docx *.xlsx *.pptx)",
            "Word Documents (*.docx)",
            "Excel Workbooks (*.xlsx)",
            "PowerPoint Presentations (*.pptx)",
            "All files (*.*)"
        ])
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.file_path = selected_files[0]
                self.file_entry.setText(self.file_path)
                self._set_status("", "green")

    def process_file(self):
        path = self.file_path
        font = self.font_entry.currentText()

        if not path:
            QMessageBox.critical(self, "Error", "Please select a file first.")
            return
        if not font:
            QMessageBox.critical(self, "Error",
                                 "Please enter a target font name.")
            return

        self._set_status("Processing...", "blue")
        self.start_button.setEnabled(False)

        self._worker = FontProcessingWorker(path, font)
        self._worker.finished.connect(self._on_processing_finished)
        self._worker.error.connect(self._on_processing_error)
        self._worker.start()

    def _on_processing_finished(self, output_path):
        self.start_button.setEnabled(True)
        self._set_status(f"Success! Saved to {output_path}", "green")
        QMessageBox.information(
            self, "Success",
            "File processed successfully and saved as: " + output_path)

    def _on_processing_error(self, message):
        self.start_button.setEnabled(True)
        self._set_status("An error occurred.", "red")
        QMessageBox.critical(
            self, "Error",
            "An error occurred during processing: " + message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FontUnifierApp()
    window.show()
    sys.exit(app.exec())
