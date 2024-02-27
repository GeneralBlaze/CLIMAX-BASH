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

# Function to extract rate figures from a PDF file
extract_rates() {
    local pdf_file="$1"
    local extracted_text=$(pdftotext "$pdf_file" -)
    
    # Find the line number where "Rate" and "SUB TOTAL NGN" are found
    local rate_start=$(echo "$extracted_text" | awk '/Rate/ {print NR; exit}')
    local subtotal_start=$(echo "$extracted_text" | awk '/SUB TOTAL NGN/ {print NR; exit}')

    # Calculate the number of lines in between "Rate" and "SUB TOTAL NGN"
    local lines_between=$((subtotal_start - rate_start))

    # Extract figures based on the number of lines between "Rate" and "SUB TOTAL NGN"
    if [ $lines_between -ge 17 ]; then
        # Extract figures from the 4th, 5th, and 6th lines after "Rate"
        local rate_figures=$(echo "$extracted_text" | awk -v start=$((rate_start + 4)) 'NR>=start && NR<=(start+2) {print}' | sort -n)
    elif [ $lines_between -eq 15 ]; then
        # Extract figures from the 4th and 5th lines after "Rate"
        local rate_figures=$(echo "$extracted_text" | awk -v start=$((rate_start + 4)) 'NR>=start && NR<=(start+1) {print}' | sort -n)
    else
        # Extract figure from the 4th line after "Rate"
        local rate_figures=$(echo "$extracted_text" | awk -v start=$((rate_start + 4)) 'NR==start {print}')
    fi

      # Split the sorted rates into separate variables
    local rate_1=$(echo "$rate_figures" | awk 'NR==1')
    local rate_2=$(echo "$rate_figures" | awk 'NR==2')
    local rate_3=$(echo "$rate_figures" | awk 'NR==3')

    # Return the sorted rates
    echo "$rate_1 $rate_2 $rate_3"

    # Print the extracted figures
    echo "$rate_figures"
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
    rates=$(extract_rates "$pdf_file")
    IFS=' ' read -r rate_1 rate_2 rate_3 <<< "$rates"
    # Add consignee, bill numbers, and rates to JSON file
    for bill_number in $bill_numbers; do
         echo "{ \"Consignee\": \"$consignee\", \"Bill Number\": \"$bill_number\", \"Rate 1\": \"$rate_1\", \"Rate 2\": \"$rate_2\", \"Rate 3\": \"$rate_3\" }," >> "$json_file"
    done
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
workbook = xlsxwriter.Workbook("$output_folder/consignee_and_bill_numbers_rates.xlsm")
worksheet = workbook.add_worksheet()

# Write column headers
worksheet.write(0, 0, "Consignee")
worksheet.write(0, 1, "Bill Number")
worksheet.write(0, 2, "Rate 1")
worksheet.write(0, 3, "Rate 2")
worksheet.write(0, 4, "Rate 3")

# Write data to Excel worksheet
row = 1
for item in data:
    worksheet.write(row, 0, item["Consignee"])
    worksheet.write(row, 1, item["Bill Number"])
    worksheet.write(row, 2, item["Rate 1"])
    worksheet.write(row, 3, item["Rate 2"])
    worksheet.write(row, 4, item["Rate 3"])
    row += 1

# Close the workbook
workbook.close()
EOF

echo "Excel spreadsheet 'consignee_and_bill_numbers.xlsx' created successfully and saved in '$output_folder' folder."
