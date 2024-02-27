#!/bin/bash

# Function to extract consignee information from a PDF file
extract_consignee() {
    local pdf_file="$1"
    local extracted_text=$(pdftotext "$pdf_file" -)
    local consignee=$(echo "$extracted_text" | awk '/Consignee:/ { getline; getline; print }')
    echo "$consignee"
}

# Function to extract bill numbers from a PDF file
extract_bill_numbers() {
    local pdf_file="$1"
    local extracted_text=$(pdftotext "$pdf_file" -)
    local bill_numbers=$(echo "$extracted_text" | awk '/BL No/ { print $NF }')
    echo "$bill_numbers"
}

# Get the current directory
current_dir=$(pwd)

# Get the list of PDF files in the current directory
pdf_files=("$current_dir"/*.pdf)

# Create a JSON file with the list of PDF files
json_file="pdf_files.json"
echo '["'${pdf_files[@]}'"]' > "$json_file"

# Process each PDF file
for pdf_file in "${pdf_files[@]}"; do
    echo "Processing file: $pdf_file"
    consignee=$(extract_consignee "$pdf_file")
    bill_numbers=$(extract_bill_numbers "$pdf_file")
    echo "Consignee: $consignee"
    echo "Bill Numbers: $bill_numbers"
    echo ""
done

echo "PDF file names saved in $json_file"
