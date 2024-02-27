#!/bin/bash

# Check if a PDF file is provided as an argument
if [ $# -ne 1 ]; then
    echo "Usage: $0 <pdf_file>"
    exit 1
fi

# Extract text from the PDF file
extracted_text=$(pdftotext "$1" -)

# Search for "Consignee:" and extract the text two lines after it
consignee=$(echo "$extracted_text" | awk '/Consignee:/ { getline; getline; print }')

# Search for "BL No" and extract the numbers after it
bill_numbers=$(echo "$extracted_text" | awk '/BL No/ { print $NF }')

# Print the extracted information
echo "Consignee: $consignee"
echo "Bill Numbers: $bill_numbers"
