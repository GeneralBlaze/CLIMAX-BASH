#!/bin/bash

# Path to the folder containing PDF files
pdf_folder="/Users/princewill/Downloads/Invoices and Recipts DCBTL/01-nov:29-nov/Reconciled 01-nov:29-nov 2023"

# Function to normalize consignee names
normalize_consignee_name() {
    local name="$1"
    # Remove non-alphanumeric characters and replace with a common separator (e.g., underscore)
    normalized_name=$(echo "$name" | tr -cd '[:alnum:]' | tr '[:upper:]' '[:lower:]' | tr ' ' '_')
    echo "$normalized_name"
}

# Function to merge similar folders
merge_folders() {
    local folder="$1"
    local first_9_chars="${folder:0:9}"
    local folders_to_merge=("$pdf_folder"/"$first_9_chars"*)
    local merged_folder="$pdf_folder/$first_9_chars"
    
    # Create the merged folder if it doesn't exist
    mkdir -p "$merged_folder"

    # Move files from similar folders to the merged folder
    for similar_folder in "${folders_to_merge[@]}"; do
        if [[ "$similar_folder" != "$merged_folder" ]]; then
            mv "$similar_folder"/* "$merged_folder"
            rmdir "$similar_folder"
        fi
    done
}

# Create a function to extract consignee names from PDF files
extract_consignee_name() {
    local pdf_file="$1"
    local extracted_text=$(pdftotext "$pdf_file" -)
    if [[ "$pdf_file" == *"RECEIPT"* ]]; then
        local consignee=$(echo "$extracted_text" | sed -n '11p' | cut -c 1-15)
    else
        local consignee=$(echo "$extracted_text" | awk -F 'Customer ID' 'NF>1{print $2}' | cut -c 5-)
    fi
    # Normalize the consignee name before returning
    normalized_consignee=$(normalize_consignee_name "$consignee")
    echo "$normalized_consignee"
}

# Move PDF files to appropriate folders based on consignee name
move_pdfs() {
    # Loop through all PDF files in the source folder
    for pdf_file in "$pdf_folder"/*.pdf; do
        # Extract consignee name from the PDF file
        consignee_name=$(extract_consignee_name "$pdf_file")

        # Create the destination folder for the consignee if it doesn't exist
        mkdir -p "$pdf_folder/$consignee_name"

        # Move the PDF file to the appropriate folder
        mv "$pdf_file" "$pdf_folder/$consignee_name"
        
        # Merge similar folders
        merge_folders "$consignee_name"
    done
}

# Main function to sort PDF files by consignee
sort_pdfs() {
    move_pdfs
}

# Call the main function
sort_pdfs
