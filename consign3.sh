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

# Check if xlsxwriter module is installed
if ! python3 -c "import xlsxwriter" &> /dev/null; then
    echo "Error: Please install the xlsxwriter module for Python 3."
    exit 1
fi

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
    # Write extracted data to JSON file
    echo "{ \"Consignee\": \"$consignee\", \"Bill Numbers\": \"$bill_numbers\" }," >> "$json_file"
done

# Remove the trailing comma and close the JSON array
sed -i '' '$ s/,$//' "$json_file"
echo "]" >> "$json_file"

# Use Python script to create Excel sheet from JSON data
python3 - <<EOF
import xlsxwriter
import json

# Load data from JSON file
with open("$json_file", "r") as f:
    data = json.load(f)

# Create a new Excel workbook and worksheet
workbook = xlsxwriter.Workbook("$current_dir/consignee_and_bill_numbers.xlsx")
worksheet = workbook.add_worksheet()

# Write column headers
worksheet.write(0, 0, "Consignee")
worksheet.write(0, 1, "Bill Numbers")

# Write data to Excel worksheet
row = 1
for item in data:
    worksheet.write(row, 0, item["Consignee"])
    worksheet.write(row, 1, item["Bill Numbers"])
    row += 1

# Close the workbook
workbook.close()
EOF

echo "Excel spreadsheet 'consignee_and_bill_numbers.xlsx' created successfully."
