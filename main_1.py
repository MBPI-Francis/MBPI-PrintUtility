import sys
import os
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QRadioButton, QLineEdit, QPushButton,
    QMessageBox, QScrollArea, QSpinBox
)
from PyQt6.QtGui import QPixmap, QPageSize, QPainter, QImage
from PyQt6.QtCore import Qt, QRect
from PyQt6.QtPrintSupport import QPrinter, QPrinterInfo


class PrintDialog(QDialog):
    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.doc = None
        self.total_pages = 0

        self.setWindowTitle("Custom Print Dialog")
        self.setGeometry(200, 200, 500, 600)

        self.load_document()

        # --- Main Layout ---
        main_layout = QVBoxLayout()
        settings_layout = QHBoxLayout()
        left_panel = QVBoxLayout()
        right_panel = QVBoxLayout()

        # --- UI Elements (unchanged) ---
        self.printer_combo = QComboBox()
        self.printer_combo.addItems(QPrinterInfo.availablePrinterNames())
        left_panel.addWidget(QLabel("Select Printer:"))
        left_panel.addWidget(self.printer_combo)

        self.paper_size_combo = QComboBox()
        self.paper_size_combo.addItems(["Letter", "A4", "Legal", "A3", "A5"])
        left_panel.addWidget(QLabel("Paper Size:"))
        left_panel.addWidget(self.paper_size_combo)

        self.color_mode_combo = QComboBox()
        self.color_mode_combo.addItems(["Color", "Grayscale"])
        right_panel.addWidget(QLabel("Color Mode:"))
        right_panel.addWidget(self.color_mode_combo)

        self.copies_spinbox = QSpinBox()
        self.copies_spinbox.setMinimum(1)
        self.copies_spinbox.setValue(1)
        right_panel.addWidget(QLabel("Copies:"))
        right_panel.addWidget(self.copies_spinbox)

        settings_layout.addLayout(left_panel)
        settings_layout.addLayout(right_panel)
        main_layout.addLayout(settings_layout)

        range_layout = QHBoxLayout()
        self.all_pages_radio = QRadioButton(f"All Pages ({self.total_pages})")
        self.all_pages_radio.setChecked(True)
        self.range_radio = QRadioButton("Pages:")
        self.page_range_edit = QLineEdit()
        self.page_range_edit.setPlaceholderText("e.g., 2 or 2-5")
        self.page_range_edit.setEnabled(False)
        self.range_radio.toggled.connect(self.page_range_edit.setEnabled)

        range_layout.addWidget(self.all_pages_radio)
        range_layout.addWidget(self.range_radio)
        range_layout.addWidget(self.page_range_edit)
        main_layout.addLayout(range_layout)

        main_layout.addWidget(QLabel("Preview:"))
        self.preview_label = QLabel("Loading preview...")
        self.preview_label.setStyleSheet("border: 1px solid black;")

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.preview_label)
        main_layout.addWidget(scroll_area)

        self.print_button = QPushButton("Print")
        self.print_button.clicked.connect(self.execute_print)
        main_layout.addWidget(self.print_button)

        self.setLayout(main_layout)
        self.render_preview()

    def load_document(self):
        try:
            if self.file_path.lower().endswith('.pdf'):
                self.doc = fitz.open(self.file_path)
                self.total_pages = len(self.doc)
            else:
                raise ValueError("Unsupported file type")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load document: {e}")
            self.close()

    def render_preview(self, page_num=0):
        try:
            if self.file_path.lower().endswith('.pdf'):
                page = self.doc.load_page(page_num)
                pix = page.get_pixmap(alpha=False)
                q_image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(q_image)
                self.preview_label.setPixmap(pixmap.scaledToWidth(400, Qt.TransformationMode.SmoothTransformation))
        except Exception as e:
            self.preview_label.setText(f"Preview not available: {e}")

    def execute_print(self):
        printer = QPrinter()
        printer.setPrinterName(self.printer_combo.currentText())
        printer.setCopyCount(self.copies_spinbox.value())

        paper_size_str = self.paper_size_combo.currentText()
        paper_sizes = {
            "Letter": QPageSize.PageSizeId.Letter, "A4": QPageSize.PageSizeId.A4,
            "Legal": QPageSize.PageSizeId.Legal, "A3": QPageSize.PageSizeId.A3,
            "A5": QPageSize.PageSizeId.A5
        }
        printer.setPageSize(QPageSize(paper_sizes.get(paper_size_str, QPageSize.PageSizeId.Letter)))

        if self.color_mode_combo.currentText() == "Grayscale":
            printer.setColorMode(QPrinter.ColorMode.GrayScale)
        else:
            printer.setColorMode(QPrinter.ColorMode.Color)

        self.process_and_print(printer)

    def process_and_print(self, printer):
        if not self.file_path.lower().endswith('.pdf'):
            QMessageBox.warning(self, "Not Implemented", "Printing is only supported for PDF files.")
            return

        try:
            pages_to_print = []
            if self.all_pages_radio.isChecked():
                pages_to_print = list(range(self.total_pages))
            else:
                range_text = self.page_range_edit.text().strip()
                if '-' in range_text:
                    start, end = map(int, range_text.split('-'))
                    pages_to_print = list(range(start - 1, end))
                else:
                    pages_to_print = [int(range_text) - 1]

            painter = QPainter()
            if not painter.begin(printer):
                QMessageBox.critical(self, "Print Error", "Could not start the printer.")
                return

            PRINT_DPI = 300

            for i, page_num in enumerate(pages_to_print):
                if not (0 <= page_num < self.total_pages):
                    continue

                page = self.doc.load_page(page_num)

                # --- FIX: Use robust, direct conversion from PyMuPDF to QImage ---
                if self.color_mode_combo.currentText() == "Grayscale":
                    pix = page.get_pixmap(dpi=PRINT_DPI, colorspace=fitz.csGRAY, alpha=False)
                    # Create a QImage from raw bytes, specifying the format is 8-bit Grayscale
                    q_image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_Grayscale8)
                else:
                    pix = page.get_pixmap(dpi=PRINT_DPI, colorspace=fitz.csRGB, alpha=False)
                    # Create a QImage from raw bytes, specifying the format is 24-bit RGB
                    q_image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)

                if q_image.isNull():
                    QMessageBox.warning(self, "Conversion Error",
                                        f"Could not create a valid image for page {page_num + 1}.")
                    continue
                # --- End of fix ---

                page_rect = printer.pageRect(QPrinter.Unit.DevicePixel)
                image_size = q_image.size()

                target_size = page_rect.size().toSize()
                scaled_size = image_size.scaled(target_size, Qt.AspectRatioMode.KeepAspectRatio)

                x = (page_rect.width() - scaled_size.width()) / 2
                y = (page_rect.height() - scaled_size.height()) / 2

                target_rect = QRect(int(x), int(y), scaled_size.width(), scaled_size.height())
                painter.drawImage(target_rect, q_image)

                if i < len(pages_to_print) - 1:
                    printer.newPage()

            painter.end()
            QMessageBox.information(self, "Success", "Document sent to printer.")
            self.accept()

        except ValueError:
            QMessageBox.critical(self, "Invalid Page Range",
                                 "Please enter a valid page or page range (e.g., '5' or '5-10').")
        except Exception as e:
            QMessageBox.critical(self, "Printing Failed", f"An unexpected error occurred: {e}")
            if painter.isActive():
                painter.end()


if __name__ == '__main__':
    app = QApplication(sys.argv)

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if os.path.exists(file_path):
            dialog = PrintDialog(file_path)
            dialog.exec()
        else:
            QMessageBox.critical(None, "Error", f"File not found: {file_path}")
    else:
        QMessageBox.information(None, "Custom Printer",
                                "To use this application, right-click a supported file and choose 'Open With'.")

    sys.exit()