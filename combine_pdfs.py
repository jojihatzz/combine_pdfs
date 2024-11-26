import os
from PyPDF2 import PdfMerger

def combine_pdfs(folder_path):
    merger = PdfMerger()

    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    
    pdf_files.sort()

    for pdf in pdf_files:
        pdf_path = os.path.join(folder_path, pdf)
        merger.append(pdf_path)

    output_file = os.path.join(folder_path, 'combined_pdfs.pdf')

    merger.write(output_file)
    merger.close()

    print(f"Combined PDF created: {output_file}")

if __name__ == "__main__":
    folder_path = input("Please enter the folder path containing PDF files: ").strip()
    
    if os.path.isdir(folder_path):
        combine_pdfs(folder_path)
    else:
        print("The provided path is not a valid directory.")
