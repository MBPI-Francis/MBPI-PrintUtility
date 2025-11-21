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
        self.preview_pages = []  # List of page numbers to show in preview
        self.current_preview_index = 0

        self.setWindowTitle("Advanced Print Dialog")
        self.setGeometry(150, 150, 550, 700)  # Slightly larger window

        self.load_document()

        # --- Main Layout ---
        main_layout = QVBoxLayout()
        settings_layout = QHBoxLayout()
        left_panel = QVBoxLayout()
        right_panel = QVBoxLayout()

        # --- Printer Selection ---
        self.printer_combo = QComboBox()
        self.printer_combo.addItems(QPrinterInfo.availablePrinterNames())
        self.printer_combo.currentTextChanged.connect(self.update_dpi_list)  # Connect to update DPI
        left_panel.addWidget(QLabel("Select Printer:"))
        left_panel.addWidget(self.printer_combo)

        # --- Paper Size ---
        self.paper_size_combo = QComboBox()
        self.paper_size_combo.addItems(["Letter", "A4", "Legal", "A3", "A5"])
        left_panel.addWidget(QLabel("Paper Size:"))
        left_panel.addWidget(self.paper_size_combo)

        # --- MODIFICATION: Dynamic DPI ComboBox ---
        self.dpi_combo = QComboBox()
        right_panel.addWidget(QLabel("Print Quality (DPI):"))
        right_panel.addWidget(self.dpi_combo)
        self.update_dpi_list()  # Initial population of DPI list

        # --- Color Mode ComboBox ---
        self.color_mode_combo = QComboBox()
        self.color_mode_combo.addItems(["Color", "Grayscale"])
        right_panel.addWidget(QLabel("Color Mode:"))
        right_panel.addWidget(self.color_mode_combo)

        # --- Copies counter ---
        self.copies_spinbox = QSpinBox()
        self.copies_spinbox.setMinimum(1)
        self.copies_spinbox.setValue(1)
        right_panel.addWidget(QLabel("Copies:"))
        right_panel.addWidget(self.copies_spinbox)

        settings_layout.addLayout(left_panel)
        settings_layout.addLayout(right_panel)
        main_layout.addLayout(settings_layout)

        # --- Page Range ---
        range_layout = QHBoxLayout()
        self.all_pages_radio = QRadioButton(f"All Pages ({self.total_pages})")
        self.all_pages_radio.setChecked(True)
        self.range_radio = QRadioButton("Pages:")
        self.page_range_edit = QLineEdit()
        self.page_range_edit.setPlaceholderText("e.g., 3 or 3-5")
        self.page_range_edit.setEnabled(False)
        self.range_radio.toggled.connect(self.page_range_edit.setEnabled)

        # Connect signals to update the preview
        self.all_pages_radio.toggled.connect(self.update_preview_range)
        self.page_range_edit.textChanged.connect(self.update_preview_range)

        range_layout.addWidget(self.all_pages_radio)
        range_layout.addWidget(self.range_radio)
        range_layout.addWidget(self.page_range_edit)
        main_layout.addLayout(range_layout)

        # --- MODIFICATION: Interactive Document Preview ---
        main_layout.addWidget(QLabel("Document Preview:"))
        self.preview_label = QLabel("Loading preview...")
        self.preview_label.setStyleSheet("border: 1px solid black;")

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.preview_label)
        main_layout.addWidget(scroll_area, 1)  # Give scroll area more space

        preview_nav_layout = QHBoxLayout()
        self.prev_page_button = QPushButton("<< Previous Page")
        self.next_page_button = QPushButton("Next Page >>")
        self.page_info_label = QLabel("Page: 1 of 1")
        self.prev_page_button.clicked.connect(self.show_prev_page)
        self.next_page_button.clicked.connect(self.show_next_page)

        preview_nav_layout.addWidget(self.prev_page_button)
        preview_nav_layout.addStretch()
        preview_nav_layout.addWidget(self.page_info_label)
        preview_nav_layout.addStretch()
        preview_nav_layout.addWidget(self.next_page_button)
        main_layout.addLayout(preview_nav_layout)

        # --- Print Button ---
        self.print_button = QPushButton("Print")
        self.print_button.clicked.connect(self.execute_print)
        main_layout.addWidget(self.print_button)

        self.setLayout(main_layout)
        self.update_preview_range()  # Initial setup of preview

    def update_dpi_list(self):
        """Populates the DPI combobox with resolutions supported by the selected printer."""
        printer_name = self.printer_combo.currentText()
        printer_info = QPrinterInfo(QPrinterInfo.printerInfo(printer_name))

        supported_resolutions = printer_info.supportedResolutions()
        self.dpi_combo.clear()

        if supported_resolutions:
            # Add resolutions as strings to the combobox
            self.dpi_combo.addItems([str(dpi) for dpi in supported_resolutions])
            # Set a sensible default if possible
            if 300 in supported_resolutions:
                self.dpi_combo.setCurrentText("300")
            elif 600 in supported_resolutions:
                self.dpi_combo.setCurrentText("600")
        else:
            # Fallback if the driver doesn't report resolutions
            self.dpi_combo.addItems(["75", "96", "150", "300", "600"])
            self.dpi_combo.setCurrentText("300")

    def update_preview_range(self):
        """Parses the page range selection and updates the preview navigation."""
        self.preview_pages = []
        try:
            if self.all_pages_radio.isChecked() or not self.page_range_edit.text():
                self.preview_pages = list(range(self.total_pages))
            else:
                range_text = self.page_range_edit.text().strip()
                # Parse ranges like "3-5"
                if '-' in range_text:
                    start, end = map(int, range_text.split('-'))
                    self.preview_pages = list(range(start - 1, end))
                # Parse single pages like "3"
                else:
                    self.preview_pages = [int(range_text) - 1]

            # Clamp values to be within the document's bounds
            self.preview_pages = [p for p in self.preview_pages if 0 <= p < self.total_pages]

        except ValueError:
            # If text is invalid (e.g., "abc"), show nothing
            self.preview_pages = []

        self.current_preview_index = 0
        self.render_current_preview_page()

    def render_current_preview_page(self):
        """Renders the currently selected page in the preview area."""
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
            self.preview_label.setPixmap(pixmap.scaledToWidth(450, Qt.TransformationMode.SmoothTransformation))
        except Exception as e:
            self.preview_label.setText(f"Error rendering page:\n{e}")

        # Update navigation UI
        self.page_info_label.setText(f"Page: {page_num_to_render + 1} of {self.total_pages}")
        self.prev_page_button.setEnabled(self.current_preview_index > 0)
        self.next_page_button.setEnabled(self.current_preview_index < len(self.preview_pages) - 1)

    def show_prev_page(self):
        """Navigates to the previous page in the preview range."""
        if self.current_preview_index > 0:
            self.current_preview_index -= 1
            self.render_current_preview_page()

    def show_next_page(self):
        """Navigates to the next page in the preview range."""
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

        # Set resolution from the DPI combobox
        try:
            resolution = int(self.dpi_combo.currentText())
            printer.setResolution(resolution)
        except (ValueError, TypeError):
            QMessageBox.warning(self, "Invalid DPI", "Could not set print resolution. Using default.")

        if self.color_mode_combo.currentText() == "Grayscale":
            printer.setColorMode(QPrinter.ColorMode.GrayScale)
        else:
            printer.setColorMode(QPrinter.ColorMode.Color)

        self.process_and_print(printer)

    def process_and_print(self, printer):
        # The list of pages to print is now the same as the preview list
        pages_to_print = self.preview_pages
        if not pages_to_print:
            QMessageBox.warning(self, "No Pages Selected", "There are no valid pages selected to print.")
            return

        painter = QPainter()
        if not painter.begin(printer):
            QMessageBox.critical(self, "Print Error", "Could not start the printer.")
            return

        # Use the selected DPI for rendering the page pixmap
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
            QMessageBox.information(self, "Success", "Document sent to printer.")
            self.accept()

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