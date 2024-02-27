#!/bin/bash

# Function to extract consignee information from a PDF file
extract_consignee() {
    local pdf_file="$1"
    local extracted_text=$(pdftotext "$pdf_file" -)
    local consignee=$(echo "$extracted_text" | awk -F 'Customer ID' 'NF>1{print $2}' | cut -c 5-)
    echo "$consignee"
}

# Function to extract bill numbers from a PDF file
extract_bill_numbers() {
    local pdf_file="$1"
    local extracted_text=$(pdftotext "$pdf_file" -)
    local bill_numbers=$(echo "$extracted_text" | awk '/BL No/ { print $NF }')
    echo "$bill_numbers"
}

# Check if xlsxwriter module is installed
if ! python3 -c "import xlsxwriter" &> /dev/null; then
    echo "Error: Please install the xlsxwriter module for Python 3."
    exit 1
fi

# Create output folder if it doesn't exist
output_folder="output"
mkdir -p "$output_folder"

# Get the current directory
current_dir=$(pwd)

# Get the list of PDF files in the current directory with "Invoice" in their filenames
pdf_files=("$current_dir"/*Invoice*.pdf)

# Create a JSON file to store extracted data
json_file="extracted_data.json"
echo "[" > "$json_file"

# Process each PDF file
for pdf_file in "${pdf_files[@]}"; do
    consignee=$(extract_consignee "$pdf_file")
    bill_numbers=$(extract_bill_numbers "$pdf_file")
    # Add consignee and bill numbers to JSON file
    for bill_number in $bill_numbers; do
        echo "{ \"Consignee\": \"$consignee\", \"Bill Number\": \"$bill_number\" }," >> "$json_file"
    done
done

# Remove the trailing comma and close the JSON array
sed -i '' '$ s/,$//' "$json_file"
echo "]" >> "$json_file"

# Use Python script to remove duplicate bill numbers and create Excel sheet from JSON data
python3 - <<EOF
import xlsxwriter
import json

# Load data from JSON file
with open("$json_file", "r") as f:
    data = json.load(f)

# Remove duplicate bill numbers
unique_data = []
seen_bill_numbers = set()
for item in data:
    if item["Bill Number"] not in seen_bill_numbers:
        unique_data.append(item)
        seen_bill_numbers.add(item["Bill Number"])

# Create a new Excel workbook and worksheet
workbook = xlsxwriter.Workbook("$output_folder/consignee_and_bill_numbers.xlsm")
worksheet = workbook.add_worksheet()

# Write column headers
worksheet.write(0, 0, "Consignee")
worksheet.write(0, 1, "Bill Number")

# Write data to Excel worksheet
row = 1
for item in unique_data:
    worksheet.write(row, 0, item["Consignee"])
    worksheet.write(row, 1, item["Bill Number"])
    row += 1

# Close the workbook
workbook.close()
EOF

echo "Excel spreadsheet 'consignee_and_bill_numbers.xlsm' created successfully and saved in '$output_folder' folder."
