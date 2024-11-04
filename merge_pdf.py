import os
from PyPDF2 import PdfMerger
import glob

# Define the input and output folders
pdf_folder = 'output2'
# Merge all PDFs
groups = ['1','2','3','4','5']
for g in groups:
    merger = PdfMerger()
    files = glob.glob(os.path.join(pdf_folder, f'*_{g}_*.pdf'))
    print(files)
    sorted_files = sorted(files, key=lambda x: int(os.path.basename(x).split('_')[0]))
    for filename in sorted_files:
        if filename.endswith('.pdf'):
            pdf_path = filename
            merger.append(pdf_path)

    # Write out the merged PDF
    merged_pdf_path = f'output3/merged_{g}.pdf'
    merger.write(merged_pdf_path)
    merger.close()
