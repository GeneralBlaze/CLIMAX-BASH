#!/bin/bash

# Directory containing the files
DIRECTORY="/Users/princewill/Downloads/test"

# Loop through all files in the directory
for file in "$DIRECTORY"/*; do
    # Extract the filename without the path
    filename=$(basename "$file")
    
    # Check if the filename contains 'cover'
    if [[ -f "$file" && "$filename" == *"_merged"* ]]; then
        # Create the new filename by replacing 'cover' with 'd'
        new_filename="${filename//_merged/}"
        new_file_path="$DIRECTORY/$new_filename"
        
        # Rename the file
        mv "$file" "$new_file_path"
        
        echo "Renamed $file to $new_file_path"
    fi
done

# #!/bin/bash

# # Directory containing the files
# DIRECTORY="/Users/princewill/Downloads/test"

# # Loop through all PDF files in the directory
# for file in "$DIRECTORY"/*.pdf; do
#     # Extract the filename without the path and extension
#     filename=$(basename "$file" .pdf)
    
#     # Create the new filename by adding 'a' before the '.pdf' extension
#     new_filename="${filename}c.pdf"
#     new_file_path="$DIRECTORY/$new_filename"
    
#     # Rename the file
#     mv "$file" "$new_file_path"
    
#     echo "Renamed $file to $new_file_path"
# done


# #!/bin/bash

# # Directory containing the files
# DIRECTORY="/Users/princewill/Downloads/test"

# # Loop through all files in the directory
# for file in "$DIRECTORY"/*; do
#     # Check if it is a file and does not already end with '.pdf'
#     if [[ -f "$file" && "$file" != *.pdf ]]; then
#         # Add '.pdf' to the end of the filename
#         new_file_path="${file}.pdf"
        
#         # Rename the file
#         mv "$file" "$new_file_path"
        
#         echo "Renamed $file to $new_file_path"
#     fi
# done