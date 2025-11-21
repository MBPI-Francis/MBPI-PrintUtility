import sys
import os
import fitz  # PyMuPDF
from PyQt6.QtPrintSupport import QPrinter
from PyQt6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QRadioButton, QLineEdit, QPushButton,
    QMessageBox, QScrollArea, QSpinBox, QWidget, QGroupBox
)
from PyQt6.QtGui import QPixmap, QPageSize, QPainter, QImage
from PyQt6.QtCore import Qt, QRect


class PrintDialog(QDialog):
    def __init__(self, file_path, parent=None):
        super().__init__(parent)

        self.setWindowFlags(
            self.windowFlags() | Qt.WindowType.WindowMinimizeButtonHint | Qt.WindowType.WindowMaximizeButtonHint)

        self.file_path = file_path
        self.doc = None
        self.total_pages = 0
        self.preview_pages = []
        self.current_preview_index = 0

        self.setWindowTitle("Professional Print Utility")
        self.setGeometry(150, 150, 900, 700)

        self.load_document()
        self.init_ui()

        self.update_dpi_list()
        self.update_preview_range()

    def init_ui(self):
        """Creates and arranges all UI widgets."""
        main_layout = QVBoxLayout(self)
        top_bar_layout = QHBoxLayout()
        content_layout = QHBoxLayout()

        self.print_button = QPushButton("Print Document")
        self.print_button.setObjectName("printButton")
        self.print_button.clicked.connect(self.execute_print)
        top_bar_layout.addStretch()
        top_bar_layout.addWidget(self.print_button)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        controls_groupbox = QGroupBox("Print Settings")
        controls_layout = QVBoxLayout(controls_groupbox)

        self.printer_combo = QComboBox()
        self.printer_combo.currentTextChanged.connect(self.update_dpi_list)
        controls_layout.addWidget(QLabel("Select Printer:"))
        controls_layout.addWidget(self.printer_combo)

        self.dpi_combo = QComboBox()
        controls_layout.addWidget(QLabel("Print Quality (DPI):"))
        controls_layout.addWidget(self.dpi_combo)

        self.color_mode_combo = QComboBox()
        self.color_mode_combo.addItems(["Color", "Grayscale"])
        controls_layout.addWidget(QLabel("Color Mode:"))
        controls_layout.addWidget(self.color_mode_combo)

        self.paper_size_combo = QComboBox()
        self.paper_size_combo.addItems(["Letter", "A4", "Legal", "A3", "A5"])
        controls_layout.addWidget(QLabel("Paper Size:"))
        controls_layout.addWidget(self.paper_size_combo)

        self.copies_spinbox = QSpinBox()
        self.copies_spinbox.setMinimum(1)
        self.copies_spinbox.setValue(1)
        controls_layout.addWidget(QLabel("Copies:"))
        controls_layout.addWidget(self.copies_spinbox)

        range_groupbox = QGroupBox("Page Range")
        range_layout = QVBoxLayout(range_groupbox)
        self.all_pages_radio = QRadioButton(f"All Pages ({self.total_pages})")
        self.all_pages_radio.setChecked(True)  # Fixed typo here
        self.range_radio = QRadioButton("Custom Range:")
        self.page_range_edit = QLineEdit()
        self.page_range_edit.setPlaceholderText("e.g., 3 or 3-5")
        self.page_range_edit.setEnabled(False)
        self.range_radio.toggled.connect(self.page_range_edit.setEnabled)
        self.all_pages_radio.toggled.connect(self.update_preview_range)
        self.page_range_edit.textChanged.connect(self.update_preview_range)

        range_layout.addWidget(self.all_pages_radio)
        range_layout.addWidget(self.range_radio)
        range_layout.addWidget(self.page_range_edit)

        left_layout.addWidget(controls_groupbox)
        left_layout.addWidget(range_groupbox)
        left_layout.addStretch()

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        preview_groupbox = QGroupBox("Document Preview")
        preview_layout = QVBoxLayout(preview_groupbox)

        self.preview_label = QLabel("Loading preview...")
        self.preview_label.setObjectName("previewLabel")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- MODIFICATION: Make scroll_area an instance attribute ---
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.preview_label)

        preview_nav_layout = QHBoxLayout()
        self.prev_page_button = QPushButton("<< Previous")
        self.next_page_button = QPushButton("Next >>")
        self.page_info_label = QLabel("Page: 1 of 1")
        self.prev_page_button.clicked.connect(self.show_prev_page)
        self.next_page_button.clicked.connect(self.show_next_page)

        preview_nav_layout.addWidget(self.prev_page_button)
        preview_nav_layout.addStretch()
        preview_nav_layout.addWidget(self.page_info_label)
        preview_nav_layout.addStretch()
        preview_nav_layout.addWidget(self.next_page_button)

        preview_layout.addWidget(self.scroll_area)  # Use the instance attribute
        preview_layout.addLayout(preview_nav_layout)
        right_layout.addWidget(preview_groupbox)

        content_layout.addWidget(left_widget, 1)
        content_layout.addWidget(right_widget, 2)

        main_layout.addLayout(top_bar_layout)
        main_layout.addLayout(content_layout)

        from PyQt6.QtPrintSupport import QPrinterInfo
        self.printer_combo.addItems(QPrinterInfo.availablePrinterNames())

    def execute_print(self):
        loading_dialog = QDialog(self)
        loading_dialog.setModal(True)
        loading_dialog.setWindowTitle("Processing...")
        loading_dialog.setLayout(QVBoxLayout())
        loading_dialog.layout().addWidget(QLabel("Sending document to printer, please wait..."))
        loading_dialog.setFixedSize(300, 100)
        loading_dialog.show()
        QApplication.processEvents()

        try:
            from PyQt6.QtPrintSupport import QPrinter
            printer = QPrinter()
            printer.setPrinterName(self.printer_combo.currentText())
            printer.setCopyCount(self.copies_spinbox.value())

            paper_size_str = self.paper_size_combo.currentText()
            paper_sizes = {
                "Letter": QPageSize.PageSizeId.Letter, "A4": QPageSize.PageSizeId.A4,
                "Legal": QPageSize.PageSizeId.Legal, "A3": QPageSize.PageSizeId.A3, "A5": QPageSize.PageSizeId.A5
            }
            printer.setPageSize(QPageSize(paper_sizes.get(paper_size_str, QPageSize.PageSizeId.Letter)))

            try:
                printer.setResolution(int(self.dpi_combo.currentText()))
            except (ValueError, TypeError):
                pass

            if self.color_mode_combo.currentText() == "Grayscale":
                printer.setColorMode(QPrinter.ColorMode.GrayScale)
            else:
                printer.setColorMode(QPrinter.ColorMode.Color)

            self.process_and_print(printer)

        finally:
            loading_dialog.close()

    def process_and_print(self, printer):
        pages_to_print = self.preview_pages
        if not pages_to_print:
            QMessageBox.warning(self, "No Pages Selected", "There are no valid pages selected to print.")
            return

        painter = QPainter()
        if not painter.begin(printer):
            QMessageBox.critical(self, "Print Error", "Could not start the printer. Check connection and drivers.")
            return

        print_dpi = printer.resolution()

        try:
            for i, page_num in enumerate(pages_to_print):
                page = self.doc.load_page(page_num)

                if self.color_mode_combo.currentText() == "Grayscale":
                    pix = page.get_pixmap(dpi=print_dpi, colorspace=fitz.csGRAY, alpha=False)
                    q_image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_Grayscale8)
                else:
                    pix = page.get_pixmap(dpi=print_dpi, colorspace=fitz.csRGB, alpha=False)
                    q_image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)

                if q_image.isNull():
                    print(f"Warning: Could not create a valid image for page {page_num + 1}.")
                    continue

                page_rect = printer.pageRect(QPrinter.Unit.DevicePixel)
                target_size = page_rect.size().toSize()
                scaled_size = q_image.size().scaled(target_size, Qt.AspectRatioMode.KeepAspectRatio)

                x = (page_rect.width() - scaled_size.width()) / 2
                y = (page_rect.height() - scaled_size.height()) / 2

                target_rect = QRect(int(x), int(y), scaled_size.width(), scaled_size.height())
                painter.drawImage(target_rect, q_image)

                if i < len(pages_to_print) - 1:
                    printer.newPage()

            painter.end()
            QMessageBox.information(self, "Success", "Document has been sent to the printer.")

        except Exception as e:
            QMessageBox.critical(self, "Printing Failed", f"An unexpected error occurred: {e}")
            if painter.isActive():
                painter.end()

    def update_dpi_list(self):
        from PyQt6.QtPrintSupport import QPrinterInfo
        printer_name = self.printer_combo.currentText()
        if not printer_name: return
        printer_info = QPrinterInfo(QPrinterInfo.printerInfo(printer_name))
        supported_resolutions = printer_info.supportedResolutions()
        self.dpi_combo.clear()
        if supported_resolutions:
            self.dpi_combo.addItems([str(dpi) for dpi in supported_resolutions])
            if 300 in supported_resolutions:
                self.dpi_combo.setCurrentText("300")
            elif 600 in supported_resolutions:
                self.dpi_combo.setCurrentText("600")
        else:
            self.dpi_combo.addItems(["75", "96", "150", "300", "600"])
            self.dpi_combo.setCurrentText("300")

    def update_preview_range(self):
        self.preview_pages = []
        try:
            if self.all_pages_radio.isChecked() or not self.page_range_edit.text():
                self.preview_pages = list(range(self.total_pages))
            else:
                range_text = self.page_range_edit.text().strip()
                if '-' in range_text:
                    start, end = map(int, range_text.split('-'))
                    self.preview_pages = list(range(start - 1, end))
                else:
                    self.preview_pages = [int(range_text) - 1]
            self.preview_pages = [p for p in self.preview_pages if 0 <= p < self.total_pages]
        except ValueError:
            self.preview_pages = []
        self.current_preview_index = 0
        self.render_current_preview_page()

    def render_current_preview_page(self):
        if not self.preview_pages:
            self.preview_label.setText("No valid pages selected.")
            self.page_info_label.setText("Page: -")
            self.prev_page_button.setEnabled(False)
            self.next_page_button.setEnabled(False)
            return

        page_num_to_render = self.preview_pages[self.current_preview_index]

        try:
            page = self.doc.load_page(page_num_to_render)
            pix = page.get_pixmap(alpha=False)
            q_image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(q_image)

            available_size = self.scroll_area.viewport().size()
            available_size.setWidth(available_size.width() - 5)
            available_size.setHeight(available_size.height() - 5)

            scaled_pixmap = pixmap.scaled(
                available_size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )

            self.preview_label.setPixmap(scaled_pixmap)

        except Exception as e:
            self.preview_label.setText(f"Error rendering page:\n{e}")

        self.page_info_label.setText(f"Page: {page_num_to_render + 1} of {self.total_pages}")
        self.prev_page_button.setEnabled(self.current_preview_index > 0)
        self.next_page_button.setEnabled(self.current_preview_index < len(self.preview_pages) - 1)

    def show_prev_page(self):
        if self.current_preview_index > 0:
            self.current_preview_index -= 1
            self.render_current_preview_page()

    def show_next_page(self):
        if self.current_preview_index < len(self.preview_pages) - 1:
            self.current_preview_index += 1
            self.render_current_preview_page()

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

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.render_current_preview_page()


if __name__ == '__main__':
    app = QApplication(sys.argv)

    try:
        with open("style.css", "r") as f:
            app.setStyleSheet(f.read())
    except FileNotFoundError:
        print("Warning: style.css not found. Using default styles.")

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