import os
import ocrmypdf

input_dir = r'C:\Users\ddavis\CODE 2\Contract Automation\auto\pdfs'
output_dir = r'C:\Users\ddavis\CODE 2\Contract Automation\auto\pdfs\After'

for pdf_file in os.listdir(input_dir):
    if pdf_file.endswith('.pdf'):
        input_path = os.path.join(input_dir, pdf_file)
        output_path = os.path.join(output_dir, pdf_file)
        try:
            ocrmypdf.ocr(input_path, output_path, deskew=True)
            print(f'Successfully processed {pdf_file}')
        except Exception as e:
            print(f'Failed to process {pdf_file}: {e}')
