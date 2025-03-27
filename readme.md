# Asset Label Generator

A Python utility to generate printable asset labels with QR codes from Excel data.

## Overview

This tool helps IT departments and inventory managers create professional asset labels from existing Excel data. It generates a print-ready PDF with labels containing:

- Asset ID
- QR code (containing the asset ID)
- Company name

Perfect for inventory tracking, IT asset management, and equipment labeling.

## Features

- Reads asset IDs directly from Excel spreadsheets
- Generates QR codes for each asset ID
- Creates a professional PDF formatted for label sheets
- Customizable label size and layout
- Configurable company name on each label
- Ready to print on standard Avery label sheets

## Requirements

- Python 3.6+
- Required libraries:
  - pandas
  - openpyxl
  - qrcode
  - pillow
  - reportlab

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/infynnity/asset-label-generator.git
   ```

2. Install required dependencies:
   ```
   pip install pandas openpyxl qrcode pillow reportlab
   ```

## Usage

1. Place your Excel file in the same directory as the script
2. Update the configuration section in the script:
   ```python
   excel_file = "your_asset_data.xlsx"  # Your Excel file name
   id_column = "Asset-ID"               # Column containing asset IDs
   company_name = "Your Company Name"   # Your company name
   ```

3. Run the script:
   ```
   python asset-label-generator.py
   ```

4. Print the generated `asset_labels.pdf` on label sheets

## Customizing Label Size

The default configuration works with standard 2×5 label sheets (10 labels per page), but you can adjust the parameters to match your specific label paper:

```python
labels_per_row = 2    # Number of labels side by side
rows_per_page = 5     # Number of rows per page
label_width = 90      # Width in mm
label_height = 40     # Height in mm
```

## Example

Input Excel with asset IDs:

| Asset-ID | Device | Serial Number |
|----------|--------|---------------|
| IT-00001 | Laptop | SN123456789   |
| IT-00002 | Monitor| MN987654321   |
| ...      | ...    | ...           |

Output: PDF with labels containing the asset ID, QR code, and company name.

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.