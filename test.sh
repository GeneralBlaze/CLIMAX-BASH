#!/bin/bash

# Check if the necessary tool is installed
if ! command -v pdftotext &> /dev/null; then
    echo "Error: pdftotext is not installed. Please install poppler-utils."
    exit 1
fi

# Check if a PDF file is provided as an argument
if [ $# -ne 1 ]; then
    echo "Usage: $0 <pdf_file>"
    exit 1
fi

# Extract text from the PDF file
pdftotext "$1" - > "$1.txt"
