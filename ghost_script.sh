#!/bin/bash

# Check if Ghostscript is installed
if ! command -v gs &> /dev/null; then
    echo "Error: Ghostscript is not installed. Please install Ghostscript to use this script."
    exit 1
fi

# Check if a directory is provided as an argument
if [ $# -ne 1 ]; then
    echo "Usage: $0 <directory>"
    exit 1
fi

directory="$1"

# Iterate over all PDF files in the directory
for input_pdf in "$directory"/*.pdf; do
    output_pdf="${input_pdf%.pdf}_compressed.pdf"

    # Compress the PDF using Ghostscript
    gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/screen -dNOPAUSE -dQUIET -dBATCH -sOutputFile="$output_pdf" "$input_pdf"

    # Check if compression was successful
    if [ $? -eq 0 ]; then
        echo "Compression successful. Compressed PDF saved as: $output_pdf"
    else
        echo "Error: Compression failed for file: $input_pdf"
    fi
done