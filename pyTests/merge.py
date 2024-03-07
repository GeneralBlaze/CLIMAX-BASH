import os
import subprocess

# Get the current directory
pdf_directory = os.getcwd()

# Dictionary to keep track of merged files
merged_files = {}

# Loop through each PDF file in the directory
for pdf_file in os.listdir(pdf_directory):
    if pdf_file.endswith(".pdf"):
        # Extract the filename without the directory path
        filename = os.path.basename(pdf_file)
        
        # Extract the prefix from the filename
        prefix = filename.split()[0]

        # Check if this prefix has been merged already
        if prefix not in merged_files:
            # Create the merged file path
            merged_file = os.path.join(pdf_directory, f"{prefix}merged.pdf")

            # Use pdftk-java to merge PDFs with the same prefix
            subprocess.run(["/usr/local/opt/pdftk-java/bin/pdftk", *[f for f in os.listdir(pdf_directory) if f.startswith(prefix)], "cat", "output", merged_file])

            # Add the merged file to the dictionary
            merged_files[prefix] = merged_file

            print(f"Merged PDFs with prefix {prefix} to {merged_file}")
