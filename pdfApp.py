import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QHBoxLayout, QDialog, QMainWindow, QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QLabel, QRadioButton, QLineEdit
from PyQt5.QtGui import QPalette, QPixmap, QBrush, QPainter, QPen, QImage, QColor
from PyQt5.QtCore import Qt, QPoint
from PyPDF2 import PdfMerger
from PyQt5.QtWidgets import QLabel

import pandas as pd
from pdf2docx import Converter
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PIL import Image
from reportlab.pdfgen import canvas
import os
import aspose.pdf as ap
from PyQt5.QtWidgets import QFileDialog, QMessageBox

import pdfplumber
import fitz  # PyMuPDF
#Correct one 

class PDFApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Utility App")
        self.setFixedSize(1000, 1200)  # Set fixed size for the window
        self.initUI()

    def initUI(self):
        image_path = PDFApp.resource_path("background.png")
        # Main layout
        layout = QVBoxLayout()

        # Background image

        pixmap = QPixmap(image_path)
        palette = QPalette()
        palette.setBrush(QPalette.Background, QBrush(pixmap))
        self.setPalette(palette)

        # Vertical layout for buttons
        button_layout = QVBoxLayout()
        button_layout.addSpacing(200)
       # Tenaris logo
       # Tenaris Logo (Center and enlarge it)
        logo_path = PDFApp.resource_path("fatimah.png")
        logo_label = QLabel(self)
        logo_pixmap = QPixmap(logo_path).scaled(350, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation)  # Increased logo size
        logo_label.setPixmap(logo_pixmap)

        # Center the logo horizontally and set its vertical position
        logo_x = (self.width() - logo_pixmap.width()) // 2  # Centers the logo horizontally
        logo_y = 200  # Adjusted vertical position
        logo_label.setGeometry(logo_x, logo_y, logo_pixmap.width(), logo_pixmap.height())
        logo_label.raise_()  # Bring the logo to the front layer

        
       
     
   
        # Merge PDFs Button
        self.merge_btn = QPushButton("Merge PDFs")
        self.setButtonStyle(self.merge_btn)
        self.merge_btn.clicked.connect(self.merge_pdfs)
        self.merge_btn.setFixedWidth(400)
        button_layout.addWidget(self.merge_btn, alignment=Qt.AlignHCenter)
        button_layout.addSpacing(20)

        # Convert PDF Button
        self.convert_btn = QPushButton("Convert PDF to Excel")
        self.setButtonStyle(self.convert_btn)
        self.convert_btn.clicked.connect(self.convert_pdf)
        self.convert_btn.setFixedWidth(400)
        button_layout.addWidget(self.convert_btn, alignment=Qt.AlignHCenter)
        button_layout.addSpacing(20)

        # Convert jpeg Button
        self.convert_btn = QPushButton("Convert JPEG to PDF")
        self.setButtonStyle(self.convert_btn)
        self.convert_btn.clicked.connect(self. merge_images_to_pdf)
        self.convert_btn.setFixedWidth(400)
        button_layout.addWidget(self.convert_btn, alignment=Qt.AlignHCenter)
        button_layout.addSpacing(20)

        # Signature Button
        self.signature_btn = QPushButton("Add Signature to PDF")
        self.setButtonStyle(self.signature_btn)
        self.signature_btn.clicked.connect(self.open_pdf_for_signature)
        self.signature_btn.setFixedWidth(400)
        button_layout.addWidget(self.signature_btn, alignment=Qt.AlignHCenter)

        # # Add the bottom label
        # # Add the bottom label after the buttons
        # self.bottom_label = QLabel("Merge PDFs, convert to Excel/JPEG, and add signatures—all in one app")
        # self.bottom_label.setStyleSheet("color: black; font-size: 30px;")
        # self.bottom_label.setAlignment(Qt.AlignCenter)

        # layout.addWidget(self.bottom_label, alignment=Qt.AlignBottom)


        # Add button layout to main layout
        layout.addLayout(button_layout)
        self.setLayout(layout)

    def setButtonStyle(self, button):
        button.setStyleSheet(""" 
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                                            stop:0 darkcyan, stop:1 teal);
                color: white;
                padding: 10px;
                font-size: 30px;
                border-radius: 20px;
            }
            QPushButton:hover {
                background: orange;
                color: black;
            }
        """)

    def merge_pdfs(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Select PDFs to Merge", "", "PDF Files (*.pdf)", options=options)
        if files:
            save_path, _ = QFileDialog.getSaveFileName(self, "Save Merged PDF", "", "PDF Files (*.pdf)")
            if save_path:
                doc = fitz.open()  # Create a new PDF document
                for pdf in files:
                    with fitz.open(pdf) as pdf_doc:
                        doc.insert_pdf(pdf_doc)  # Merge PDFs
                doc.save(save_path)  # Save the merged PDF
                doc.close()
                QMessageBox.information(self, "Success", "PDFs merged successfully!")

    #convert pdf to excel
    def convert_pdf(self):
        options = QFileDialog.Options()
        file, _ = QFileDialog.getOpenFileName(self, "Select PDF to Convert", "", "PDF Files (*.pdf)", options=options)
        if file:
            convert_type, ok = QFileDialog.getSaveFileName(self, "Save As", "", "Excel Files (*.xlsx);;Word Files (*.docx)")
            if ok:
                if convert_type.endswith(".xlsx"):
                    try:
                        # Convert PDF to Excel using Aspose.PDF
                        document = ap.Document(file)
                        save_option = ap.ExcelSaveOptions()
                        document.save(convert_type, save_option)
                        QMessageBox.information(self, "Success", "PDF converted to Excel successfully!")
                    except Exception as e:
                        QMessageBox.warning(self, "Error", f"Failed to convert PDF to Excel. Error: {e}")
                elif convert_type.endswith(".docx"):
                    try:
                        # Convert PDF to Word using Aspose.PDF
                        document = ap.Document(file)
                        save_option = ap.DocSaveOptions()
                        save_option.format = ap.DocSaveOptions.DocFormat.DocX
                        document.save(convert_type, save_option)
                        QMessageBox.information(self, "Success", "PDF converted to Word successfully!")
                    except Exception as e:
                        QMessageBox.warning(self, "Error", f"Failed to convert PDF to Word. Error: {e}")
    #convert pdf to jpeg

    def merge_images_to_pdf(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Select Images to Convert to PDF", "", "Image Files (*.jpg *.png *.jpeg)", options=options)
        
        if files:
            output_pdf, _ = QFileDialog.getSaveFileName(self, "Save As", "", "PDF Files (*.pdf)")
            if output_pdf:
                try:
                    # Create a PDF canvas
                    c = canvas.Canvas(output_pdf)
                    
                    for fs in files:
                        # Open the image using Pillow
                        img = Image.open(fs)
                        
                        # Get image dimensions
                        width, height = img.size
                        
                        # Convert pixels to points (1 pixel = 0.75 points)
                        width_pt = width * 0.75
                        height_pt = height * 0.75
                        
                        # Set the PDF page size to match the image size
                        c.setPageSize((width_pt, height_pt))
                        
                        # Draw the image on the PDF
                        c.drawImage(fs, 0, 0, width_pt, height_pt)
                        
                        # End the current page
                        c.showPage()
                    
                    # Save the PDF
                    c.save()
                    
                    QMessageBox.information(self, "Success", "Images merged into PDF successfully!")
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Failed to merge images into PDF. Error: {str(e)}")




    def open_pdf_for_signature(self):
        self.signatureApp = PdfSignatureApp()
        self.signatureApp.show()
    def resource_path(relative_path):
        """ Get the absolute path to the resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        full_path = os.path.join(base_path, relative_path)
        print(f"Loading resource from: {full_path}")  # Debug output
        return full_path
    resource_path("background.png")
     

class SignatureWidget(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.setRenderHint(QPainter.Antialiasing)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.image = QImage(200, 100, QImage.Format_ARGB32)
        self.image.fill(Qt.white)
        self.pixmapItem = QGraphicsPixmapItem()
        self.scene.addItem(self.pixmapItem)
        self.drawing = False
        self.lastPoint = QPoint()
        self.setFixedSize(220, 120)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drawing = True
            self.lastPoint = self.mapToScene(event.pos()).toPoint()

    def mouseMoveEvent(self, event):
        if self.drawing:
            currentPoint = self.mapToScene(event.pos()).toPoint()
            painter = QPainter(self.image)
            pen_color = QColor(0, 0, 0)
            pen_width = 2
            pen = QPen(pen_color, pen_width, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
            painter.setPen(pen)
            painter.drawLine(self.lastPoint, currentPoint)
            self.lastPoint = currentPoint
            self.pixmapItem.setPixmap(QPixmap.fromImage(self.image))
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drawing = False

    def saveImage(self):
        # Define the path where the signature will be saved
        temp_dir = os.path.join(os.path.expanduser("~"), "pdfTool_temp")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        filename = os.path.join(temp_dir, "signature.png")
        self.image.save(filename, "PNG")
        return filename  # Return the path of the saved signature
    def get_temp_directory(self):
        # Define a custom directory for temporary files
        temp_dir = os.path.join(os.path.expanduser("~"), "pdfTool_temp")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        return temp_dir


    def clear_signature(self):
        self.image.fill(Qt.white)
        self.pixmapItem.setPixmap(QPixmap.fromImage(self.image))
        self.update()

class PdfSignatureApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Signature Window")
        self.setGeometry(100, 100, 900, 700)
        self.initUI()
        self.pdf_file = None

    def initUI(self):
        self.loadPdfButton = QPushButton("Load PDF", self)
        self.loadPdfButton.clicked.connect(self.load_pdf)

        self.savePdfButton = QPushButton("Save PDF with Signature", self)
        self.savePdfButton.clicked.connect(self.save_pdf)

        self.clearSignatureButton = QPushButton("Clear Signature", self)
        self.clearSignatureButton.clicked.connect(self.clear_signature)

        self.signatureWidget = SignatureWidget()
        self.pdfViewer = PdfViewer()

        self.nextPageButton = QPushButton("Next Page", self)
        self.nextPageButton.clicked.connect(self.pdfViewer.next_page)

        self.prevPageButton = QPushButton("Previous Page", self)
        self.prevPageButton.clicked.connect(self.pdfViewer.previous_page)

        self.firstPageRadio = QRadioButton("First Page")
        self.lastPageRadio = QRadioButton("Last Page")

        self.dateInput = QLineEdit()
        self.dateInput.setPlaceholderText("Date (e.g., YYYY-MM-DD)")
        self.nameInput = QLineEdit()
        self.nameInput.setPlaceholderText("Name")

        signature_layout = QHBoxLayout()
        signature_layout.addWidget(QLabel("Signature:"))
        signature_layout.addWidget(self.signatureWidget)
        signature_layout.addWidget(self.dateInput)
        signature_layout.addWidget(self.nameInput)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.loadPdfButton)
        button_layout.addWidget(self.prevPageButton)
        button_layout.addWidget(self.nextPageButton)
        button_layout.addWidget(self.clearSignatureButton)
        button_layout.addWidget(self.savePdfButton)

        layout = QVBoxLayout()
        layout.addLayout(signature_layout)
        layout.addLayout(button_layout)
        layout.addWidget(QLabel("PDF Viewer:"))
        layout.addWidget(self.pdfViewer)
        layout.addWidget(self.firstPageRadio)
        layout.addWidget(self.lastPageRadio)

        self.firstPageRadio.setChecked(True)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def load_pdf(self):
        options = QFileDialog.Options()
        self.pdf_file, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)", options=options)
        if self.pdf_file:
            self.pdfViewer.load_pdf(self.pdf_file)
    # def get_temp_directory(self):
    #     # Define a custom directory for temporary files
    #     temp_dir = os.path.join(os.path.expanduser("~"), "pdfTool_temp")
    #     if not os.path.exists(temp_dir):
    #         os.makedirs(temp_dir)
    #     return temp_dir

    def save_pdf(self):
        if not self.pdf_file:
            QMessageBox.warning(self, "Warning", "No PDF file loaded.")
            return

        # Save the signature image without passing any argument
        signature_image_path = self.signatureWidget.saveImage()  # Call without arguments

        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
        if save_path:
            # Open the original PDF with PyMuPDF
            pdf_document = fitz.open(self.pdf_file)

            # Determine page index based on radio button selection
            if self.firstPageRadio.isChecked():
                page_index = 0  # Insert before the first page
                pdf_document.insert_page(page_index)  # Create a new page
            else:
                page_index = pdf_document.page_count  # Insert after the last page
                pdf_document.insert_page(page_index)  # Create a new page

            # Get the newly created page for signature
            page = pdf_document[page_index]

            # Use the path returned by saveImage
            rect = fitz.Rect(50, 50, 250, 150)  # Adjust size and position
            page.insert_image(rect, filename=signature_image_path)  # Use the correct path

            # Insert the name and date
            date_text = self.dateInput.text()
            name_text = self.nameInput.text()
            page.insert_text((50, 200), f"Date: {date_text}", fontsize=12, color=(0, 0, 0))
            page.insert_text((50, 220), f"Name: {name_text}", fontsize=12, color=(0, 0, 0))

            # Save the modified PDF
            pdf_document.save(save_path)
            pdf_document.close()
            QMessageBox.information(self, "Success", "Signature added and PDF saved successfully.")


    def clear_signature(self):
        self.signatureWidget.clear_signature()

class PdfViewer(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.setRenderHint(QPainter.Antialiasing)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.pdf_document = None
        self.current_page = 0
        self.setFixedSize(600, 800)

    def load_pdf(self, pdf_path):
        self.pdf_document = fitz.open(pdf_path)
        self.current_page = 0
        self.show_page(self.current_page)

    def show_page(self, page_number):
        if self.pdf_document:
            page = self.pdf_document[page_number]
            pix = page.get_pixmap()
            image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            self.scene.clear()
            self.scene.addPixmap(QPixmap.fromImage(image))
            self.setScene(self.scene)

    def next_page(self):
        if self.pdf_document and self.current_page < self.pdf_document.page_count - 1:
            self.current_page += 1
            self.show_page(self.current_page)

    def previous_page(self):
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.show_page(self.current_page)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFApp()
    window.show()
    sys.exit(app.exec_())
