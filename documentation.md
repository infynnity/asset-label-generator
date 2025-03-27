# Asset Label Generator - Documentation

## Overview

This Python program creates a PDF with labels for your assets. Each label contains:
- The Asset ID from your Excel spreadsheet
- A QR code containing the Asset ID
- Your company name

The labels are formatted for printing on standard label sheets and can be printed with any office printer.

## Prerequisites

Before you can use the program, you need to install the following Python libraries:

```
pip install pandas openpyxl qrcode pillow reportlab
```

## Installation

1. Save the program code in a file named `asset_label_generator.py`
2. Place your Excel file with the Asset IDs in the same directory

## Configuration

In the lower part of the program, you'll find a section that you can adapt to your needs:

```python
if __name__ == "__main__":
    # Use the current directory (where the script is located)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Adjust these values as needed
    excel_file = os.path.join(current_dir, "your_asset_data.xlsx")  # Your Excel file
    id_column = "Asset-ID"  # Column with IDs (from the screenshot)
    company_name = "Your Company Name"  # Replace with your company name
    
    # Generate PDF with asset labels
    generate_asset_labels(
        excel_file=excel_file,
        id_column=id_column,
        company_name=company_name,
        output_pdf=os.path.join(current_dir, "asset_labels.pdf"),
        # Adjust these settings to match your label sticker size
        labels_per_row=2,        # Labels per row
        rows_per_page=5,         # Rows per page
        label_width=90,          # in mm (standard width for Avery labels)
        label_height=40,         # in mm (standard height for Avery labels)
    )
```

Here you need to adjust:
- `excel_file`: Name of your Excel file
- `id_column`: Name of the column containing the Asset IDs
- `company_name`: Name of your company
- Label format settings (depending on the label paper you use)

## Usage

1. **Prepare Excel file**
   - Make sure your Excel file has a column named "Asset-ID" (or the name you specified)
   - The file should be in the same directory as the Python script

2. **Run the program**
   - Open a command line/terminal
   - Navigate to the directory with the script
   - Run the command: `python asset_label_generator.py`
   - The program will create a PDF file named "asset_labels.pdf" in the same directory

3. **Print the labels**
   - Open the created PDF file
   - Load label paper into your printer
   - Print the PDF file without scaling (100% size, "Actual size")

## Adjusting the Label Size

If you want to use a different label format, adjust the following parameters:

- `labels_per_row`: Number of labels side by side on the sheet
- `rows_per_page`: Number of label rows per sheet
- `label_width`: Width of each label in millimeters
- `label_height`: Height of each label in millimeters

For common Avery label formats:
- 2x5 layout (10 per page): e.g., Avery L7666
- 3x7 layout (21 per page): e.g., Avery L7160

## Troubleshooting

**Problem**: The program cannot find the Excel file.
**Solution**: Make sure the file name is correct and the file is in the same directory as the script.

**Problem**: The column "Asset-ID" is not found.
**Solution**: Check if the column name in your Excel file exactly matches the value of `id_column`.

**Problem**: The labels are not properly aligned on the label paper.
**Solution**: Adjust the values for `label_width` and `label_height` according to your label paper.