#!/bin/bash

# Get the current directory
pdf_directory=$(pwd)

# Dictionary to keep track of merged files
declare -A merged_files

# Loop through each PDF file in the directory
for pdf_file in "$pdf_directory"/*.pdf; do
    if [[ -f "$pdf_file" ]]; then
        # Extract the filename without the directory path
        filename=$(basename "$pdf_file")
        
        # Extract the prefix from the filename
        prefix=$(echo "$filename" | awk -F' ' '{print $1}')

        # Check if this prefix has been merged already
        if [[ -z "${merged_files[$prefix]}" ]]; then
            # Create the merged file path
            merged_file="$pdf_directory/${prefix}merged.pdf"

            # Use pdftk-java to merge PDFs with the same prefix
            /usr/local/opt/pdftk-java/bin/pdftk $(ls "$pdf_directory" | grep "^$prefix") cat output "$merged_file"

            # Add the merged file to the dictionary
            merged_files[$prefix]=$merged_file

            echo "Merged PDFs with prefix $prefix to $merged_file"
        fi
    fi
done
