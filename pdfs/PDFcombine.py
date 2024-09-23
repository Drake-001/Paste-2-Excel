import os
from PyPDF2 import PdfMerger
import re

# Specify the folder containing your PDFs
folder_path = r'C:\Users\ddavis\CODE 2\pdfs\Prose'

# Custom sorting function to sort by the numeric part of the filenames
def sort_key(filename):
    # Extract the numeric part of the filename (assuming format like '1.pdf', '2.pdf', etc.)
    numbers = re.findall(r'\d+', filename)
    return int(numbers[0]) if numbers else 0

# Get all the PDF files from the folder and sort them by the numeric part of the filename
pdf_files = sorted([f for f in os.listdir(folder_path) if f.endswith('.pdf')], key=sort_key)

# Create a PdfMerger object
merger = PdfMerger()

# Loop through each PDF file and append it to the merger
for pdf in pdf_files:
    pdf_path = os.path.join(folder_path, pdf)
    merger.append(pdf_path)

# Write the merged PDF to a new file
output_path = os.path.join(folder_path, r'C:\Users\ddavis\CODE 2\pdfs\Prose\prose2\merged.pdf')
merger.write(output_path)
merger.close()

print(f"All PDFs have been merged into {output_path}")
