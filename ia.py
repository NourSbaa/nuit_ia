import os
import pandas as pd
from PyPDF2 import PdfReader
from pdf2image import convert_from_path

def extract_images_from_pdf(pdf_path, output_folder):
    """
    Extract images from a PDF file and save them in the output folder.

    Args:
        pdf_path (str): Path to the PDF file.
        output_folder (str): Path to the output folder where images will be saved.

    Returns:
        list: List of paths to the extracted images.
    """
    images = []
    with open(pdf_path, 'rb') as f:
        reader = PdfReader(f)
        for i, page in enumerate(reader.pages):
            # Utilisez le chemin vers le fichier PDF comme premier argument
            # de convert_from_path et spécifiez le dossier de sortie avec output_folder
            images += convert_from_path(pdf_path, output_folder=output_folder, first_page=i+1, last_page=i+1)
    return images

def main():
    # Chemin du fichier PDF
    pdf_path = '/content/PDF.pdf'
    # Dossier de sortie pour les images extraites
    output_folder = 'extracted_images'

    # Créer le dossier de sortie s'il n'existe pas
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Extraire les images du PDF
    extracted_images = extract_images_from_pdf(pdf_path, output_folder)

    # Convertir les chemins des images en noms de fichiers
    image_files = [os.path.join(output_folder, f) for f in os.listdir(output_folder) if f.endswith('.jpg')]

    # Créer un DataFrame Pandas avec les noms des fichiers d'images
    df = pd.DataFrame({'Filename': image_files})

    # Enregistrer le DataFrame dans un fichier Excel
    excel_file = 'extracted_images.xlsx'
    df.to_excel(excel_file, index=False)
    print(f"Extracted images saved to {excel_file}")

if __name__ == "__main__":
    main()
