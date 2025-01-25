# pip install PyPDF2 comtypes

import os
import sys
import comtypes.client
from PyPDF2 import PdfMerger

def convert_ppt_to_pdf(ppt_path, pdf_path):
    """
    Converts a PowerPoint file to PDF using Microsoft PowerPoint.

    :param ppt_path: Path to the PowerPoint file.
    :param pdf_path: Path where the PDF will be saved.
    """
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    try:
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        presentation.SaveAs(pdf_path, 32)  # 32 for ppSaveAsPDF
        presentation.Close()
        print(f"Converted '{ppt_path}' to '{pdf_path}'.")
    except Exception as e:
        print(f"Failed to convert '{ppt_path}' to PDF. Error: {e}")
    finally:
        powerpoint.Quit()

def combine_pdfs(folder_path, output_filename='combined_pdfs.pdf'):
    """
    Combines all PDF files in the specified folder into a single PDF.

    :param folder_path: Path to the folder containing PDF files.
    :param output_filename: Name of the output combined PDF file.
    """
    merger = PdfMerger()

    # Gather all PDF files
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    pdf_files.sort()

    if not pdf_files:
        print("No PDF files found to combine.")
        return

    for pdf in pdf_files:
        pdf_path = os.path.join(folder_path, pdf)
        try:
            merger.append(pdf_path)
            print(f"Added '{pdf}' to the merger.")
        except Exception as e:
            print(f"Failed to add '{pdf}' to the merger. Error: {e}")

    output_file = os.path.join(folder_path, output_filename)

    try:
        merger.write(output_file)
        print(f"Combined PDF created: '{output_file}'")
    except Exception as e:
        print(f"Failed to write the combined PDF. Error: {e}")
    finally:
        merger.close()

def main():
    folder_path = input("Please enter the folder path containing PDF and PowerPoint files: ").strip()

    if not os.path.isdir(folder_path):
        print("The provided path is not a valid directory.")
        sys.exit(1)

    # Convert PowerPoint files to PDF
    ppt_extensions = ['.ppt', '.pptx']
    converted_pdfs = []

    for file in os.listdir(folder_path):
        file_lower = file.lower()
        if any(file_lower.endswith(ext) for ext in ppt_extensions):
            ppt_path = os.path.join(folder_path, file)
            pdf_filename = os.path.splitext(file)[0] + '.pdf'
            pdf_path = os.path.join(folder_path, pdf_filename)

            # Check if PDF already exists to avoid reconversion
            if not os.path.exists(pdf_path):
                convert_ppt_to_pdf(ppt_path, pdf_path)
            else:
                print(f"PDF already exists for '{ppt_path}'. Skipping conversion.")

    # Combine all PDFs
    combine_pdfs(folder_path)

if __name__ == "__main__":
    main()
