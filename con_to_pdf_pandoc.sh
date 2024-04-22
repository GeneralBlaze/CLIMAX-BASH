#!/bin/bash

# Add pandoc and pdflatex to the PATH
export PATH="/usr/local/bin:/usr/local/texlive/2024basic/bin/universal-darwin:$PATH"


# Function to convert Word documents to PDF
convert_to_pdf() {
    local dir=$1
    local doc_file doc_name pdf_name

    echo "Starting conversion process..."

    # Convert all Word documents in the selected directory to PDF files
    for doc_file in "$dir"/*.doc*; do
        if [[ -f "$doc_file" ]]; then
            doc_name=$(basename "$doc_file")
            pdf_name="${doc_name%.*}.pdf"
            pandoc "$doc_file" -o "$dir/$pdf_name"
            if [[ $? -eq 0 ]]; then
                echo "Converted $doc_name to PDF: $pdf_name"
            else
                echo "Error converting $doc_name to PDF."
                return 1
            fi
        fi
    done

    echo "Conversion process completed."
}

# Function to check if any PDF files were created
check_pdf_creation() {
    local dir=$1
    local pdf_files

    # Check if any PDF files were created
    pdf_files=$(find "$dir" -type f -name "*.pdf")
    if [[ -n "$pdf_files" ]]; then
        echo "Conversion complete. PDF files saved in: $dir"
    else
        echo "No PDF files were created."
        return 1
    fi
}

# Prompt the user to select a directory using AppleScript
selected_dir=$(osascript -e 'tell application "Finder" to choose folder with prompt "Select a directory"')

# Convert the AppleScript-style path to a Unix-style path
selected_dir=$(echo "$selected_dir" | sed 's/alias //' | sed 's/:/\//g' | sed 's/^Macintosh HD//')

# Check if the user canceled the selection
if [[ -z "$selected_dir" ]]; then
    echo "No directory selected. Exiting."
    exit 1
fi

# Convert Word documents to PDF and check the result
if ! convert_to_pdf "$selected_dir"; then
    echo "An error occurred during the conversion process. Exiting."
    exit 1
fi

# Check if any PDF files were created and check the result
if ! check_pdf_creation "$selected_dir"; then
    echo "No PDF files were created. Exiting."
    exit 1
fi