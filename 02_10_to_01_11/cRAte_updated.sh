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

# Get the current date and time
current_datetime=$(date +"%Y-%m-%d_%H-%M-%S")

# Get the current directory
current_dir=$(pwd)

# Get the list of PDF files in the current directory with "Invoice" in their filenames
pdf_files=("$current_dir"/*Invoice.pdf)

# Create a JSON file to store extracted data
json_file="extracted_data.json"
echo "[" > "$json_file"

# Set to store processed bill numbers
processed_bill_numbers=()

# Process each PDF file
for pdf_file in "${pdf_files[@]}"; do
    consignee=$(extract_consignee "$pdf_file")
    bill_numbers=$(extract_bill_numbers "$pdf_file")
    rates=$(extract_rates "$pdf_file")
    IFS=' ' read -r rate_1 rate_2 rate_3 <<< "$rates"
    # Add consignee, bill numbers, and rates to JSON file
    for bill_number in $bill_numbers; do
        # Check if bill number is already processed
        if ! [[ " ${processed_bill_numbers[@]} " =~ " $bill_number " ]]; then
            echo "{ \"Consignee\": \"$consignee\", \"Bill Number\": \"$bill_number\", \"Rate 1\": \"$rate_1\", \"Rate 2\": \"$rate_2\", \"Rate 3\": \"$rate_3\" }," >> "$json_file"
            processed_bill_numbers+=("$bill_number")
        fi
    done
done

# Remove the trailing comma and close the JSON array
sed -i '' '$ s/,$//' "$json_file"
echo "]" >> "$json_file"

# Use Python script to create Excel sheet from JSON data
python3 - <<EOF
import xlsxwriter
import json
import openpyxl

# Load data from JSON file
with open("$json_file", "r") as f:
    data = json.load(f)

# Create a new Excel workbook and worksheet
workbook = xlsxwriter.Workbook("$output_folder/ConBilRate_${current_datetime}.xlsx")
worksheet = workbook.add_worksheet()

# Define a format for the title row
title_format = workbook.add_format({'bold': True, 'bg_color': '#AA6E15'})

# Write column headers
worksheet.write_row(0, 0, ["Consignee", "Bill Number", "Rate 1", "Rate 2", "Rate 3"], title_format)

# Write data to Excel worksheet
row = 1
for item in data:
    worksheet.write_row(row, 0, [item["Consignee"], item["Bill Number"], item["Rate 1"], item["Rate 2"], item["Rate 3"]])
    row += 1

# Adjust column widths based on the maximum length of the longest cell value in each column
for col_idx, col_data in enumerate(zip(*data)):
    max_length = max(len(str(cell)) for cell in col_data) + 2  # Adding padding
    worksheet.set_column(col_idx, col_idx, max_length)

# Adjust cell widths based on the length of cell values plus padding before and after
for row_idx, row_data in enumerate(data):
    for col_idx, cell_value in enumerate(row_data.values()):
        max_length = len(str(cell_value)) + 2  # Adding padding
        worksheet.set_column(col_idx, col_idx, max_length)

# Close the workbook
workbook.close()

# Load the Excel sheet again to update empty s2 and s3 rates based on matching s1 rates
wb = openpyxl.load_workbook("$output_folder/ConBilRate_${current_datetime}.xlsx")
ws = wb.active

# Step 1: Find rows with empty s2 and s3 rates and their corresponding s1 rates
empty_rates = {}
for row in ws.iter_rows(min_row=2, values_only=True):
    consignee, bill_number, rate_1, rate_2, rate_3 = row[0], row[1], row[2], row[3], row[4]
    if rate_2 is None or rate_3 is None:
        if rate_1 not in empty_rates:
            empty_rates[rate_1] = []
        empty_rates[rate_1].append((consignee, bill_number))

# Step 2: Fill empty s2 and s3 rates based on matching s1 rates
for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    consignee, bill_number, rate_1, rate_2, rate_3 = row[0], row[1], row[2], row[3], row[4]
    if rate_2 is None or rate_3 is None:
        for other_row_idx, other_row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if other_row_idx != row_idx:
                other_rate_1, other_rate_2, other_rate_3 = other_row[2], other_row[3], other_row[4]
                if rate_1 == other_rate_1:
                    if rate_2 is None:
                        ws.cell(row=row_idx, column=4, value=other_rate_2)
                    if rate_3 is None:
                        ws.cell(row=row_idx, column=5, value=other_rate_3)
                    break


# Save the updated workbook
wb.save("$output_folder/ConBil_RatesUpdated${current_datetime}_updated.xlsx")

EOF

echo "Excel spreadsheet 'ConBilRate_${current_datetime}_updated.xlsx' created successfully and saved in '$output_folder' folder."
