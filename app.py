import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader
import pandas as pd
import re
import os
import pdfplumber
from PIL import Image as PILImage  # Import the Image class from PIL
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils.dataframe import dataframe_to_rows  # Correct import statement


# Function to extract images
def extract_images(pdf_path, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    img_paths = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            images = page.images
            for j, img in enumerate(images):
                img_path = os.path.join(output_folder, f"image_page_{i+1}_img_{j+1}.png")
                with open(img_path, "wb") as f:
                    f.write(img["stream"].get_data())
                img_paths.append(img_path)
    return img_paths

# Function to resize images
def resize_image(img_path, output_folder):
    img = PILImage.open(img_path)  # Open the image with PIL
    resized_img_path = os.path.join(output_folder, f"resized_{os.path.basename(img_path)}")
    img.thumbnail((100, 100))  # Resize the image
    img.save(resized_img_path)  # Save the resized image
    return resized_img_path

# Function to extract data from PDF
def extract_data_from_pdf(pdf_path):
    output_folder = os.path.dirname(pdf_path)  # Output folder for images
    images = extract_images(pdf_path, output_folder)  # Extract images
    resized_images = [resize_image(img, output_folder) for img in images]  # Resize images

    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        num_pages = len(reader.pages)

        # Initialize lists to store extracted data
        data = {
            'Picture': [],
            "Supplier's Reference": [],
            "Supplier's Designation": [],
            'Product Range': [],
            'Colour(s)': [],
            'Measure Units': [],
            'Brand/License': [],
            'BIUB or BBD*(dd/mm/yyyy)': [],
            'Untaxed (Wine)': [],
            'Qty Available': [],
            'Wholesale Price': [],
            'Clearance Price': [],
            'Retail Price': [],
            'Packing Details': [],
            'Nb Packets / Pallet': [],
            'Number of Pallets': []
        }

        # Loop through each page
        for page_num in range(num_pages):
            page = reader.pages[page_num]
            text = page.extract_text()

            # Extracting data using regular expressions
            picture = resized_images[page_num] if page_num < len(resized_images) else None  # Use the resized image
            supplier_ref = re.search(r'(?<=\-\s)[^-\n]+', text)
            supplier_desig = re.search(r'(?<=Supplier\'s reference\n)[^\n]+', text)  # Modified regex
            product_range_default = 'Accessoires'
            colour = re.search(r'(?<=\- )[^-\n]+', text)
            measure_unit_default = 'One Size Fits most'
            brand_license_default = 'C.C'
            biub_or_bbd_default = ''
            untaxed_default = ''
            qty_avail = re.search(r'(?<=Qty available: )\d+', text)
            wholesale_price_default = ''
            clearance_price_default = ''
            retail_price_default = ''
            packing_details_default = ''
            nb_packets_pallet_default = ''

            # Extracting Number of Pallets from text after colon (':') and writing it to column J
            num_pallets_match = re.search(r'(?<=Number of Pallets: )\d+', text)
            number_of_pallets = num_pallets_match.group(0) if num_pallets_match else ''

            # Append extracted data to respective lists
            data['Picture'].append(picture)
            data["Supplier's Reference"].append(supplier_ref.group(0) if supplier_ref else '')
            data["Supplier's Designation"].append(supplier_desig.group(0) if supplier_desig else '')  # Modified condition
            data['Product Range'].append(product_range_default)
            data['Colour(s)'].append(colour.group(0) if colour else '')
            data['Measure Units'].append(measure_unit_default)
            data['Brand/License'].append(brand_license_default)
            data['BIUB or BBD*(dd/mm/yyyy)'].append(biub_or_bbd_default)
            data['Untaxed (Wine)'].append(untaxed_default)
            data['Qty Available'].append(qty_avail.group(0) if qty_avail else '')
            data['Wholesale Price'].append(wholesale_price_default)
            data['Clearance Price'].append(clearance_price_default)
            data['Retail Price'].append(retail_price_default)
            data['Packing Details'].append(packing_details_default)
            data['Nb Packets / Pallet'].append(nb_packets_pallet_default)
            data['Number of Pallets'].append(number_of_pallets)

    # Create a DataFrame using the extracted data
    df = pd.DataFrame(data)
    return df

# Function to handle button click event
def browse_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_entry.delete(0, tk.END)
    pdf_entry.insert(0, pdf_path)

def extract_to_excel():
    pdf_path = pdf_entry.get()
    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file.")
        return
    
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        return  # User canceled save operation
    
    try:
        df = extract_data_from_pdf(pdf_path)
        wb = Workbook()
        ws = wb.active

        # Insert images into Excel
        for idx, row in df.iterrows():
            if row['Picture']:
                img = PILImage.open(row['Picture'])
                img = ExcelImage(img)
                ws.add_image(img, f'A{idx+2}')

        # Write DataFrame to Excel
        for r_idx, r in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            ws.append(r)
            # Adjust row height
            ws.row_dimensions[r_idx].height = 100  # Adjust the height as needed

        # Adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2  # Adjust this factor as needed
            ws.column_dimensions[column].width = adjusted_width

        wb.save(excel_path)
        messagebox.showinfo("Success", "Extraction completed successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main window
root = tk.Tk()
root.title("PDF to Excel Converter")

# Create and place widgets
pdf_label = tk.Label(root, text="Select PDF file:")
pdf_label.grid(row=0, column=0, padx=5, pady=5)

pdf_entry = tk.Entry(root, width=50)
pdf_entry.grid(row=0, column=1, padx=5, pady=5, columnspan=2)

browse_button = tk.Button(root, text="Browse", command=browse_pdf)
browse_button.grid(row=0, column=3, padx=5, pady=5)

convert_button = tk.Button(root, text="Convert to Excel", command=extract_to_excel)
convert_button.grid(row=1, column=0, columnspan=4, padx=5, pady=5)

# Run the main event loop
root.mainloop()
