import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import pdfplumber
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image

# Function to extract images
def extract_images(pdf_path, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            images = page.images
            for j, img in enumerate(images):
                img_path = os.path.join(output_folder, f"image_page_{i+1}_img_{j+1}.png")
                with open(img_path, "wb") as f:
                    f.write(img["stream"].get_data())
    return img_path

# Function to extract text above bold text
def extract_text_above_bold(pdf_path):
    bold_text_above = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split("\n")
            for i, line in enumerate(lines):
                if re.search(r'\b\w+:\b', line):  # Check for ':' character
                    bold_text_above.append(lines[i-1])
    return bold_text_above

# Function to extract text after '-'
def extract_text_after_dash(pdf_path):
    text_after_dash = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split("\n")
            for line in lines:
                match = re.search(r'-(.*)', line)
                if match:
                    text_after_dash.append(match.group(1).strip())
    return text_after_dash

# Main function to orchestrate the extraction
def extract_data_from_pdf(pdf_path):
    output_folder = os.path.dirname(pdf_path)
    output_filename = os.path.splitext(os.path.basename(pdf_path))[0] + "_output.xlsx"
    output_path = os.path.join(output_folder, output_filename)
    
    # Extract images
    img_path = extract_images(pdf_path, output_folder)

    # Extract data for each column
    supplier_reference = extract_text_above_bold(pdf_path)
    supplier_designation = extract_text_above_bold(pdf_path)
    product_range = "Accessory"  # Default value
    colours = extract_text_after_dash(pdf_path)
    measure_units = "One size fits most"  # Default value
    brand_license = "C.C"  # Default value
    qty_available = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            matches = re.findall(r'(?<=: )\d+', text)
            qty_available.extend(matches)

    # Create and populate Excel file
    wb = Workbook()
    ws = wb.active
    ws.append(["Picture", "Supplier's reference", "Supplier's designation", "PRODUCT RANGE", "Colour(s)", "Measure units", "Brand/License", "BIUB or BBD*(dd/mm/yyyy)", "Untaxed (Wine)", "Qty available", "Wholesale Price", "ClearancePrice", "Retail price", "Packing details", "Nb packets / pallet", "Number of pallets"])

    # Add image to the first cell
    img = Image(img_path)
    ws.add_image(img, 'A2')

    # Populate other columns
    for ref, des, color, qty in zip(supplier_reference, supplier_designation, colours, qty_available):
        ws.append(["", ref, des, product_range, color, measure_units, brand_license, "", "", qty, "", "", "", "", "", ""])

    # Save the workbook
    wb.save(output_path)
    return output_path

# Function to handle file selection
def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(tk.END, file_path)

# Function to extract data and display in GUI
def extract_and_display():
    pdf_path = pdf_entry.get()
    if not pdf_path:
        messagebox.showwarning("Warning", "Please select a PDF file.")
        return
    
    output_path = extract_data_from_pdf(pdf_path)
    
    messagebox.showinfo("Info", f"Extraction complete. Output file saved as:\n{output_path}")

# Create GUI window
root = tk.Tk()
root.title("PDF Data Extraction")
root.geometry("600x400")

# Create widgets
pdf_label = tk.Label(root, text="Select PDF file:")
pdf_label.pack(pady=(20, 5))

pdf_entry = tk.Entry(root, width=50)
pdf_entry.pack(pady=(0, 5), padx=10)

browse_button = tk.Button(root, text="Browse", command=browse_pdf)
browse_button.pack()

extract_button = tk.Button(root, text="Extract and Display Data", command=extract_and_display)
extract_button.pack(pady=(10, 20))

root.mainloop()