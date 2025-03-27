import pandas as pd
import qrcode
import os
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from io import BytesIO
from PIL import Image as PILImage

def generate_asset_labels(
    excel_file, 
    id_column,
    company_name="Your Company Name", 
    output_pdf="asset_labels.pdf",
    labels_per_row=2,
    rows_per_page=5,
    label_width=90,  # in mm
    label_height=40,  # in mm
    page_size=A4
):
    """
    Generate a PDF with asset labels containing:
    - Asset ID
    - Company name
    - QR code for the Asset ID
    
    Formatted for printing on sticker paper.
    
    Parameters:
    - excel_file: Path to Excel file with asset IDs
    - id_column: Column name containing the asset IDs
    - company_name: Company name to print on labels
    - output_pdf: Output PDF file path
    - labels_per_row: Number of labels in each row
    - rows_per_page: Number of rows of labels per page
    - label_width: Width of each label in mm
    - label_height: Height of each label in mm
    - page_size: Page size (A4 or letter)
    """
    print(f"Reading Excel file: {excel_file}")
    
    # Read the Excel file
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # Check if the column exists
    if id_column not in df.columns:
        print(f"Error: Column '{id_column}' not found in the Excel file")
        print(f"Available columns: {df.columns.tolist()}")
        return
    
    # Filter out any rows with empty ID values
    df = df[df[id_column].notna()]
    asset_ids = df[id_column].astype(str).tolist()
    
    total_ids = len(asset_ids)
    if total_ids == 0:
        print("No valid asset IDs found in the file")
        return
    
    print(f"Found {total_ids} asset IDs to process")
    
    # Create PDF document
    pdf = SimpleDocTemplate(output_pdf, pagesize=page_size, 
                           leftMargin=12*mm, rightMargin=12*mm,
                           topMargin=12*mm, bottomMargin=12*mm)
    
    # Set up styles
    styles = getSampleStyleSheet()
    id_style = ParagraphStyle(
        'AssetID',
        parent=styles['Heading2'],
        fontSize=12,
        leading=14
    )
    company_style = ParagraphStyle(
        'CompanyName',
        parent=styles['Normal'],
        fontSize=8,
        leading=10
    )
    
    # Convert mm to points (reportlab uses points)
    label_width_pt = label_width * mm
    label_height_pt = label_height * mm
    
    # Function to create one label (returns a list of elements)
    def create_label(asset_id):
        # Generate QR code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=4,
        )
        qr.add_data(asset_id)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        # Convert PIL image to reportlab Image
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        buffer.seek(0)
        
        # Calculate QR code size (about 60% of label height)
        qr_size = label_height_pt * 0.6
        
        # Create components
        qr_image = Image(buffer, width=qr_size, height=qr_size)
        id_text = Paragraph(f"<b>{asset_id}</b>", id_style)
        company_text = Paragraph(company_name, company_style)
        
        # Create a table for layout
        data = [[qr_image, id_text], ["", company_text]]
        col_widths = [qr_size + 5, label_width_pt - qr_size - 5]
        row_heights = [label_height_pt * 0.7, label_height_pt * 0.3]
        
        # Create table
        table = Table(data, colWidths=col_widths, rowHeights=row_heights)
        
        # Add style to table
        table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (1, 0), (1, -1), 'LEFT'),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        
        return table
    
    # Create labels
    story = []
    
    # Calculate labels per page
    labels_per_page = labels_per_row * rows_per_page
    
    # Process all asset IDs
    for i in range(0, total_ids, labels_per_page):
        # Process one page of labels at a time
        page_ids = asset_ids[i:i+labels_per_page]
        
        # Create rows of labels
        for j in range(0, len(page_ids), labels_per_row):
            row_ids = page_ids[j:j+labels_per_row]
            
            # Create label tables for this row
            label_tables = [create_label(id_) for id_ in row_ids]
            
            # If row is not full, add empty cells
            while len(label_tables) < labels_per_row:
                label_tables.append(Spacer(label_width_pt, label_height_pt))
            
            # Create a table for this row of labels
            row_table = Table([label_tables], 
                             colWidths=[label_width_pt] * labels_per_row,
                             rowHeights=[label_height_pt])
            
            # Add spacing between cells
            row_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            
            story.append(row_table)
            story.append(Spacer(1, 5*mm))  # Add space between rows
            
            # Print progress
            progress = min(i + j + len(row_ids), total_ids)
            if progress % 100 == 0 or progress == total_ids:
                print(f"Progress: {progress}/{total_ids} labels generated")
        
        # Add page break after each page except the last one
        if i + labels_per_page < total_ids:
            story.append(Spacer(1, 0))
    
    # Build PDF
    print(f"Building PDF...")
    pdf.build(story)
    print(f"Labels PDF generated successfully at: {output_pdf}")
    print(f"Ready for printing on sticker paper with {labels_per_row}x{rows_per_page} layout")

if __name__ == "__main__":
    # Use the current directory (same as where the script is located)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Set these values according to your needs
    excel_file = os.path.join(current_dir, "your_asset_data.xlsx")  # Your Excel file
    id_column = "Asset-ID"  # Column with IDs (from screenshot)
    company_name = "Your Company Name"  # Replace with your company name
    
    # Generate PDF with asset labels
    generate_asset_labels(
        excel_file=excel_file,
        id_column=id_column,
        company_name=company_name,
        output_pdf=os.path.join(current_dir, "asset_labels.pdf"),
        # Adjust these settings to match your label sticker size
        labels_per_row=2,
        rows_per_page=5,
        label_width=90,  # in mm (standard Avery label width)
        label_height=40,  # in mm (standard Avery label height)
    )