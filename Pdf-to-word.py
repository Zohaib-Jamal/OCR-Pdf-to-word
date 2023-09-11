import fitz  # PyMuPDF
import docx
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image

# Set the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update this path

# Function to convert a PDF to a Word document using OCR
def pdf_to_word_with_ocr(pdf_file, word_file):
    # Initialize a Word document
    doc = docx.Document()

    # Open the PDF file using PyMuPDF
    pdf_document = fitz.open(pdf_file)

    # Iterate through each page in the PDF
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)

        # Get a PIL (Pillow) image from the pixmap
        image = page.get_pixmap()
        pil_image = Image.frombytes("RGB", [image.width, image.height], image.samples)

        # Save the PIL image as a temporary PNG file
        img_path = f"page_{page_num}.png"
        pil_image.save(img_path)

        # Perform OCR on the extracted image
        text = pytesseract.image_to_string(Image.open(img_path))

        # Add the extracted text to the Word document
        doc.add_paragraph(text)

    # Save the Word document
    doc.save(word_file)

    # Clean up: Remove temporary image files
    for page_num in range(pdf_document.page_count):
        img_path = f"page_{page_num}.png"
        try:
            Image.open(img_path).close()
            Image.open(img_path).unlink()
        except Exception as e:
            pass

    messagebox.showinfo("Conversion Complete", f"PDF '{pdf_file}' has been converted to Word document '{word_file}'.")

# Function to handle the "Convert" button click
def convert_pdf_to_word():
    pdf_file = pdf_file_entry.get()
    word_file = word_file_entry.get()

    if not pdf_file or not word_file:
        messagebox.showerror("Error", "Please select PDF and Word file paths.")
        return

    try:
        pdf_to_word_with_ocr(pdf_file, word_file)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Function to browse for a PDF file
def browse_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        pdf_file_entry.delete(0, tk.END)
        pdf_file_entry.insert(0, file_path)

# Function to browse for a Word file
def browse_word_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
    if file_path:
        word_file_entry.delete(0, tk.END)
        word_file_entry.insert(0, file_path)

# Create the main application window
app = tk.Tk()
app.title("PDF to Word Converter")

# Create and pack GUI elements
pdf_label = tk.Label(app, text="Select PDF File:")
pdf_label.pack()

pdf_file_entry = tk.Entry(app, width=50)
pdf_file_entry.pack()

browse_pdf_button = tk.Button(app, text="Browse", command=browse_pdf_file)
browse_pdf_button.pack()

word_label = tk.Label(app, text="Select Word File:")
word_label.pack()

word_file_entry = tk.Entry(app, width=50)
word_file_entry.pack()

browse_word_button = tk.Button(app, text="Browse", command=browse_word_file)
browse_word_button.pack()

convert_button = tk.Button(app, text="Convert", command=convert_pdf_to_word)
convert_button.pack()

# Customize the appearance of the buttons
browse_pdf_button.config(font=('Helvetica', 12))
browse_word_button.config(font=('Helvetica', 12))
convert_button.config(font=('Helvetica', 14, 'bold'))

# Customize the appearance of labels and entry fields
pdf_label.config(font=('Helvetica', 14))
word_label.config(font=('Helvetica', 14))
pdf_file_entry.config(font=('Helvetica', 12))
word_file_entry.config(font=('Helvetica', 12))

app.mainloop()
