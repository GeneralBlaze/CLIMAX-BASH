import os
from PyPDF2 import PdfMerger

# Prompt for the directory
directory = r"/Users/princewill/alx-interview/CLIMAX BASH/REFUND LETTERS"

#sort list of pdf for easy locating
sorted_list = sorted(os.listdir(directory))

# Find all .pdf files in the directory
pdf_files = [f for f in sorted_list if f.endswith("pdf")]

print(f"Found {len(pdf_files)} PDF files.")

# Initialize a PdfMerger object
merger = PdfMerger()

# Merge all the .pdf files
for filename in pdf_files:
    print(f"Processing {filename}...")
    try:
        merger.append(os.path.join(directory, filename))
        print(f"Processed {filename}.")
    except Exception as e:
        print(f"Failed to process {filename}: {e}")

# Write the merged .pdf file to disk
output_filename = os.path.join(directory, "order REFUND LETTERS.pdf")
merger.write(output_filename)
merger.close()

print(f"Merged {len(pdf_files)} PDFs into {output_filename}")