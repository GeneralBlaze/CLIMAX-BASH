#!/bin/bash

# Explicitly set environment variables
PATH=/usr/local/bin:$PATH

LOG_FILE=~/Desktop/sort_folders.log

# Function to select a folder with Finder
select_folder() {
    osascript <<EOT
        tell application "Finder"
            activate
            return POSIX path of (choose folder with prompt "PLEASE SELECT THE FOLDER WITH THE PDF FILES:")
        end tell
EOT
}

# Function to select a file with Finder
select_file() {
    osascript <<EOT
        tell application "Finder"
            activate
            return POSIX path of (choose file with prompt "PLEASE SELECT THE FILE WITH THE CONSIGNEE NAMES:")
        end tell
EOT
}

# Paths to the folder containing PDF files and the JSON file containing the list of names to compare against
# These are now selected with Finder
readonly PDF_FOLDER="$(select_folder)"
readonly JSON_FILE="$(select_file)"

# Function to normalize consignee names
normalize_consignee_name() {
    local name="$1"
    # Remove non-alphanumeric characters and replace with a common separator (e.g., underscore)
    echo "$name" | tr -cd '[:alnum:]' | tr '[:upper:]' '[:lower:]' | tr ' ' '_'
}

# Function to merge similar folders
merge_folders() {
    local folder="$1"
    local first_9_chars="${folder:0:9}"
    local folders_to_merge=("$PDF_FOLDER"/"$first_9_chars"*)
    local merged_folder="$PDF_FOLDER/$first_9_chars"
    
    # Create the merged folder if it doesn't exist
    mkdir -p "$merged_folder"

    # Move files from similar folders to the merged folder
    for similar_folder in "${folders_to_merge[@]}"; do
        if [[ "$similar_folder" != "$merged_folder" ]]; then
            mv "$similar_folder"/* "$merged_folder" && rmdir "$similar_folder"
        fi
    done
}

# Function to rename folders based on JSON file
rename_folders() {
    local names
    names=$(jq -r '.names | .[]' "$JSON_FILE")

    for name in $names; do
        local matched_folders=("$PDF_FOLDER"/*"$name"*)
        if [[ ${#matched_folders[@]} -gt 0 ]]; then
            for matched_folder in "${matched_folders[@]}"; do
                local new_name
                new_name=$(jq -r --arg name "$name" '.rename[$name]' "$JSON_FILE")
                mv "$matched_folder" "$PDF_FOLDER/$new_name"
            done
        fi
    done
}

# Function to extract consignee names from PDF files
extract_consignee_name() {
    local pdf_file="$1"
    local extracted_text
    pdftotext=/usr/local/bin/pdftotext
    extracted_text=$("$pdftotext" "$pdf_file" -)
    local consignee

    if [[ "$pdf_file" == *"RECEIPT"* ]]; then
        consignee=$(echo "$extracted_text" | sed -n '11p' | cut -c 1-15)
    else
        consignee=$(echo "$extracted_text" | awk -F 'Customer ID' 'NF>1{print $2}' | cut -c 5-)
    fi

    # Normalize the consignee name before returning
    normalize_consignee_name "$consignee"
}

# Function to move PDF files to appropriate folders based on consignee name
move_pdfs() {
    # Loop through all PDF files in the source folder
    for pdf_file in "$PDF_FOLDER"/*.pdf; do
        # Extract consignee name from the PDF file
        local consignee_name
        consignee_name=$(extract_consignee_name "$pdf_file")

        # Create the destination folder for the consignee if it doesn't exist
        mkdir -p "$PDF_FOLDER/$consignee_name"

        # Move the PDF file to the appropriate folder
        mv "$pdf_file" "$PDF_FOLDER/$consignee_name"
        
        # Merge similar folders
        merge_folders "$consignee_name"
    done
}

# Main function to sort PDF files by consignee
sort_pdfs() {
    move_pdfs
    rename_folders
}

# Call the main function
sort_pdfs