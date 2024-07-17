import os
import json
import shutil
from PyPDF2 import PdfMerger

def merge_pdfs_by_tracking_number(json_path, folder_path):
    """
    Merges PDF files based on tracking numbers. PDF files containing the tracking number in their file name
    and ending with either '.pdf' or 'pdf' (without the dot) are merged into a single PDF for each tracking number.
    Merged PDFs with 4 or more files are saved in one folder, and those with 3 or fewer files are saved in another folder.

    Parameters:
    - json_path (str): Path to the JSON file containing tracking numbers.
    - folder_path (str): Path to the folder containing PDF files to be merged.
    """
    # Load tracking numbers from JSON
    with open(json_path, 'r') as json_file:
        tracking_numbers = json.load(json_file)
        
    processed_folder_path = os.path.join(folder_path, "processed")
    if not os.path.exists(processed_folder_path):
        os.makedirs(processed_folder_path)

    merged_4_or_more_path = os.path.join(folder_path, "merged_4_or_more")
    if not os.path.exists(merged_4_or_more_path):
        os.makedirs(merged_4_or_more_path)

    merged_3_or_less_path = os.path.join(folder_path, "merged_3_or_less")
    if not os.path.exists(merged_3_or_less_path):
        os.makedirs(merged_3_or_less_path)

    for tracking_number in tracking_numbers:
        merger = PdfMerger()
        merged = False
        matching_files = [] # List to store matching PDF files

        for file_name in os.listdir(folder_path):
            if file_name.lower().endswith(('.pdf', 'pdf')) and tracking_number in file_name:
                pdf_path = os.path.join(folder_path, file_name)
                matching_files.append(pdf_path)

        # Sort the matching files before merging
        matching_files.sort()

        for pdf_path in matching_files:
            merger.append(pdf_path)
            merged = True
            # Move processed file to the 'processed' folder
            shutil.move(pdf_path, os.path.join(processed_folder_path, os.path.basename(pdf_path)))

        if merged:
            if len(matching_files) >= 4:
                output_path = os.path.join(merged_4_or_more_path, f"{tracking_number}.pdf")
            else:
                output_path = os.path.join(merged_3_or_less_path, f"{tracking_number}.pdf")
            
            merger.write(output_path)
            merger.close()
            print(f"Merged PDF saved as {output_path}")

# Example usage
json_path = r'/Users/princewill/alx-interview/CLIMAX BASH/trn.json'
folder_path = r'/Users/princewill/Downloads/test'
merge_pdfs_by_tracking_number(json_path, folder_path)