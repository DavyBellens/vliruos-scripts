#!/bin/bash

# Level
level="ICP"

# Ensure required commands are available
command -v pdftotext >/dev/null 2>&1 || { echo "pdftotext command not found. Please install it first."; exit 1; }
command -v grep >/dev/null 2>&1 || { echo "grep command not found. Please install it first."; exit 1; }

# List of institutions
institutions=("Hogeschool Gent" "Hogeschool PXL" "Hogeschool Vives" "Vives Hogeschool" "Artesis Plantijn" "Erasmushogeschool Brussel" "KU Leuven" "Odisee" "Thomas More" "UGent" "Universiteit Antwerpen" "Universiteit Gent" "Universiteit Hasselt" "Vrije Universiteit Brussel" "VUB" "Hogeschool West-Vlaanderen" "Hogeschool WestVlaanderen")

# Function to process a single file
process_file() {
    local file="$1"
    local textfile="${file%.pdf}.txt"

    # Convert PDF to text
    pdftotext "$file" "$textfile"
    if [[ ! -f "$textfile" ]]; then
        echo "Failed to convert $file to text."
        return
    fi

    # Initialize found flag
    local found=0

    # Search for each institution in the text file
    for institution in "${institutions[@]}"; do
        if grep -q "$institution" "$textfile"; then
            found=1
            mkdir -p "../extracted/$level/$institution"
            cp "$file" "../extracted/$level/$institution"
            rm "$textfile"
            break
        fi
    done

    # Handle files that didn't match any institution
    if [[ $found -eq 0 ]]; then
        echo "Failed to classify $file"
    fi

    # Remove the temporary text file
}

# Main script logic
main() {
    # Change to the data directory
    cd data || { echo "Data directory not found"; exit 1; }

    # Process each PDF file in the directory
    find . -maxdepth 1 -name '*.pdf' | while read -r file; do
        process_file "$file"
    done

    # Change back to the parent directory
    cd ..
}

# Run the main function
main
