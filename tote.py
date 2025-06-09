import streamlit as st
import pandas as pd
import os
from reportlab.lib.pagesizes import landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak, Image
from reportlab.lib.units import cm, inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.utils import ImageReader
from io import BytesIO
import subprocess
import sys
import re
import tempfile

# Define sticker dimensions - Fixed as per original code
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# Define content box dimensions - Fixed as per original code
CONTENT_BOX_WIDTH = 8 * cm
CONTENT_BOX_HEIGHT = 3 * cm

# Fixed column width proportions for the 7-box layout (no sliders)
COLUMN_WIDTH_PROPORTIONS = [1.0, 1.9, 0.8, 0.8, 0.7, 0.7, 0.8]

# Fixed content positioning
CONTENT_LEFT_OFFSET = 1.4 * cm

# Check for PIL and install if needed
try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    st.error("Installing PIL...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pillow'])
    from PIL import Image as PILImage
    PIL_AVAILABLE = True

# Check for QR code library and install if needed
try:
    import qrcode
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False
    st.error("Installing qrcode...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'qrcode'])
    import qrcode
    QR_AVAILABLE = True

# Define paragraph styles - Fixed font sizes as per original
bold_style = ParagraphStyle(name='Bold', fontName='Helvetica-Bold', fontSize=9, alignment=TA_CENTER, leading=10)
desc_style = ParagraphStyle(name='Desc', fontName='Helvetica', fontSize=7, alignment=TA_LEFT, leading=9)
qty_style = ParagraphStyle(name='Quantity', fontName='Helvetica', fontSize=8, alignment=TA_CENTER, leading=11)

def generate_qr_code(data_string):
    """Generate a QR code from the given data string"""
    try:
        # Create QR code instance
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=4,
        )
        
        # Add data
        qr.add_data(data_string)
        qr.make(fit=True)
        
        # Create QR code image
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        # Convert PIL image to bytes that reportlab can use
        img_buffer = BytesIO()
        qr_img.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        
        # Create a QR code image with fixed size
        return Image(img_buffer, width=1.5*cm, height=1.5*cm)
    except Exception as e:
        st.error(f"Error generating QR code: {e}")
        return None

def parse_location_string(location_str):
    """Parse a location string into components for table display"""
    # Initialize with empty values - 7 components to match table structure
    location_parts = [''] * 7

    if not location_str or not isinstance(location_str, str):
        return location_parts

    # Remove any extra spaces
    location_str = location_str.strip()

    # Try to parse location components
    pattern = r'([^_\s]+)'
    matches = re.findall(pattern, location_str)

    # Fill the available parts - up to 7 parts
    for i, match in enumerate(matches[:7]):
        location_parts[i] = match

    return location_parts

def generate_sticker_labels(df, progress_bar=None, status_container=None):
    """Generate sticker labels with QR code from DataFrame"""
    
    # Create a function to draw the border box around content
    def draw_border(canvas, doc):
        canvas.saveState()
        x_offset = CONTENT_LEFT_OFFSET
        y_offset = STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.8*cm
        canvas.setStrokeColor(colors.Color(0, 0, 0, alpha=0.95))
        canvas.setLineWidth(1.5)
        canvas.rect(
            x_offset,
            y_offset,
            CONTENT_BOX_WIDTH,
            CONTENT_BOX_HEIGHT
        )
        canvas.restoreState()

    # Identify columns (case-insensitive)
    original_columns = df.columns.tolist()
    df.columns = [col.upper() if isinstance(col, str) else col for col in df.columns]
    cols = df.columns.tolist()

    # Find relevant columns
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                   next((col for col in cols if col in ['PARTNO', 'PART']), cols[0]))

    desc_col = next((col for col in cols if 'DESC' in col),
                   next((col for col in cols if 'NAME' in col), cols[1] if len(cols) > 1 else part_no_col))

    # Look specifically for "QTY/BIN" column first
    qty_bin_col = next((col for col in cols if 'QTY/BIN' in col or 'QTY_BIN' in col or 'QTYBIN' in col), 
                  next((col for col in cols if 'QTY' in col and 'BIN' in col), None))
    
    # If no specific QTY/BIN column is found, fall back to general QTY column
    if not qty_bin_col:
        qty_bin_col = next((col for col in cols if 'QTY' in col),
                      next((col for col in cols if 'QUANTITY' in col), None))
  
    loc_col = next((col for col in cols if 'LOC' in col or 'POS' in col or 'LOCATION' in col),
                   cols[2] if len(cols) > 2 else desc_col)

    # Look for store location column
    store_loc_col = next((col for col in cols if 'STORE' in col and 'LOC' in col),
                      next((col for col in cols if 'STORELOCATION' in col), None))

    if status_container:
        status_container.write(f"**Using columns:**")
        status_container.write(f"- Part No: {part_no_col}")
        status_container.write(f"- Description: {desc_col}")
        status_container.write(f"- Location: {loc_col}")
        status_container.write(f"- Qty/Bin: {qty_bin_col}")
        if store_loc_col:
            status_container.write(f"- Store Location: {store_loc_col}")

    # Create temporary file for PDF output
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_path = temp_file.name
    temp_file.close()

    # Create document with minimal margins
    doc = SimpleDocTemplate(temp_path, pagesize=STICKER_PAGESIZE,
                          topMargin=0.1*cm, bottomMargin=0.1*cm,
                          leftMargin=0.1*cm, rightMargin=0.1*cm)

    all_elements = []

    # Process each row as a single sticker
    total_rows = len(df)
    for index, row in df.iterrows():
        # Update progress
        if progress_bar:
            progress_bar.progress((index + 1) / total_rows)
        
        if status_container:
            status_container.write(f"Creating sticker {index+1} of {total_rows}")
        
        elements = []

        # Extract data
        part_no = str(row[part_no_col])
        desc = str(row[desc_col])
        
        # Extract QTY/BIN properly
        qty_bin = ""
        if qty_bin_col and qty_bin_col in row and pd.notna(row[qty_bin_col]):
            qty_bin = str(row[qty_bin_col])
            
        location_str = str(row[loc_col]) if loc_col and loc_col in row else ""
        store_location = str(row[store_loc_col]) if store_loc_col and store_loc_col in row else ""
        location_parts = parse_location_string(location_str)

        # Generate QR code with part information
        qr_data = f"Part No: {part_no}\nDescription: {desc}\nLocation: {location_str}\n"
        qr_data += f"Store Location: {store_location}\nQTY/BIN: {qty_bin}"
        
        qr_image = generate_qr_code(qr_data)
        
        # Define row heights - Fixed sizes as per original
        header_row_height = 0.6*cm
        desc_row_height = 0.8*cm
        qty_row_height = 0.5*cm
        location_row_height = 0.5*cm

        # Fixed dimensions
        qr_width = 1.5*cm  
        main_content_width = CONTENT_BOX_WIDTH - qr_width

        # Main table data - Fixed column widths
        header_col_width = main_content_width * 0.22
        content_col_width = main_content_width * 0.71
        
        main_table_data = [
            ["Part No", Paragraph(f"{part_no}", bold_style)],
            ["Desc", Paragraph(desc[:30] + "..." if len(desc) > 30 else desc, desc_style)],
            ["Q/B", Paragraph(str(qty_bin), qty_style)]
        ]

        # Create main table with fixed column widths
        main_table = Table(main_table_data,
                         colWidths=[header_col_width, content_col_width],
                         rowHeights=[header_row_height, desc_row_height, qty_row_height])

        main_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.0, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (0, -1), 8),
        ]))

        # Store Location section - Fixed layout
        store_loc_label = Paragraph("S.LOC", ParagraphStyle(
            name='StoreLoc', fontName='Helvetica-Bold', fontSize=7, alignment=TA_CENTER
        ))

        # Fixed width for the inner columns
        inner_table_width = content_col_width
        
        # Calculate column widths based on fixed proportions 
        total_proportion = sum(COLUMN_WIDTH_PROPORTIONS)
        inner_col_widths = [w * inner_table_width / total_proportion for w in COLUMN_WIDTH_PROPORTIONS]

        # Use store_location if available, otherwise use empty values
        store_loc_values = parse_location_string(store_location) if store_location else ["", "", "", "", "", "", ""]

        store_loc_inner_table = Table(
            [store_loc_values],
            colWidths=inner_col_widths,
            rowHeights=[location_row_height]
        )

        store_loc_inner_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.0, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
        ]))

        store_loc_table = Table(
            [[store_loc_label, store_loc_inner_table]],
            colWidths=[header_col_width, inner_table_width],
            rowHeights=[location_row_height]
        )

        store_loc_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.0, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        # Line Location section - Fixed layout
        line_loc_label = Paragraph("L.LOC", ParagraphStyle(
            name='LineLoc', fontName='Helvetica-Bold', fontSize=7, alignment=TA_CENTER
        ))
        
        # Create the inner table for location_parts using the same fixed widths
        line_loc_inner_table = Table(
            [location_parts],
            colWidths=inner_col_widths,
            rowHeights=[location_row_height]
        )
        
        line_loc_inner_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.0, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8)
        ]))
        
        # Wrap the label and the inner table in a containing table
        line_loc_table = Table(
            [[line_loc_label, line_loc_inner_table]],
            colWidths=[header_col_width, inner_table_width],
            rowHeights=[location_row_height]
        )

        line_loc_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.0, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        # Create main content table (combining all the content tables vertically)
        total_main_height = header_row_height + desc_row_height + qty_row_height
        
        main_content_table = Table(
            [[main_table], [store_loc_table], [line_loc_table]],
            colWidths=[main_content_width],
            rowHeights=[total_main_height, location_row_height, location_row_height]
        )

        main_content_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        # QR code table - Fixed positioning
        if qr_image:
            qr_table = Table(
                [[qr_image]],
                colWidths=[qr_width], 
                rowHeights=[CONTENT_BOX_HEIGHT]
            )
        else:
            qr_table = Table(
                [[Paragraph("QR", ParagraphStyle(
                    name='QRPlaceholder', fontName='Helvetica-Bold', fontSize=10, alignment=TA_CENTER
                ))]],
                colWidths=[qr_width],
                rowHeights=[CONTENT_BOX_HEIGHT]
            )

        qr_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        # Final layout with fixed dimensions
        final_table = Table(
            [[main_content_table, qr_table]],
            colWidths=[main_content_width, qr_width],
            rowHeights=[CONTENT_BOX_HEIGHT]
        )

        final_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        # Fixed spacer
        elements.append(Spacer(1, 0.3*cm))
        elements.append(final_table)

        # Add all elements for this sticker to the document
        all_elements.extend(elements)

        # Add page break after each sticker (except the last one)
        if index < len(df) - 1:
            all_elements.append(PageBreak())

    # Build the document
    try:
        doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
        if status_container:
            status_container.success("PDF generated successfully!")
        return temp_path
    except Exception as e:
        if status_container:
            status_container.error(f"Error building PDF: {e}")
        return None

def main():
    st.set_page_config(
        page_title="Tote Label Generator",
        page_icon="🏷️",
        layout="wide"
    )
    
    st.title("🏷️Tote Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )

    st.markdown("---")
    
    # Sidebar for file upload
    with st.sidebar:
        st.header("📁 File Upload")
        uploaded_file = st.file_uploader(
            "Choose Excel or CSV file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your Excel or CSV file containing product data"
        )
        
        if uploaded_file:
            st.success(f"File uploaded: {uploaded_file.name}")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if uploaded_file is not None:
            try:
                # Read the file
                if uploaded_file.name.lower().endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                st.subheader("📊 Data Preview")
                st.write(f"**Total rows:** {len(df)}")
                st.write(f"**Columns:** {', '.join(df.columns.tolist())}")
                
                # Show preview of data
                st.dataframe(df.head(10), use_container_width=True)
                
                # Generate button
                if st.button("🚀 Generate Tote Labels", type="primary", use_container_width=True):
                    with st.spinner("Generating sticker labels..."):
                        # Create containers for progress and status
                        progress_bar = st.progress(0)
                        status_container = st.empty()
                        
                        # Generate the PDF
                        pdf_path = generate_sticker_labels(df, progress_bar, status_container)
                        
                        if pdf_path:
                            # Read the generated PDF
                            with open(pdf_path, 'rb') as pdf_file:
                                pdf_data = pdf_file.read()
                            
                            # Clean up temporary file
                            os.unlink(pdf_path)
                            
                            # Download button
                            st.download_button(
                                label="📥 Download PDF",
                                data=pdf_data,
                                file_name=f"{uploaded_file.name.split('.')[0]}_sticker_labels.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                            
                            st.success("✅ Tote labels generated successfully!")
                        else:
                            st.error("❌ Failed to generate tote labels")
                            
            except Exception as e:
                st.error(f"Error reading file: {str(e)}")
        else:
            st.info("👈 Please upload an Excel or CSV file to get started")
    
    with col2:
        st.subheader("ℹ️ Information")
        
        st.markdown("""
        **Fixed Layout Specifications:**
        - **Sticker Size:** 10cm × 15cm
        - **Content Box:** 8cm × 3cm
        - **QR Code:** 1.5cm × 1.5cm
        - **7-Box Layout:** Fixed proportions
        - **L.LOC & S.LOC:** Fixed sizes (no sliders)
        """)
        
        st.markdown("""
        **Expected Columns:**
        - Part Number (PART, PARTNO, etc.)
        - Description (DESC, NAME, etc.)
        - Quantity/Bin (QTY/BIN, QTY, etc.)
        - Location (LOC, LOCATION, POS, etc.)
        - Store Location (STORE LOC, etc.)
        """)
        
        st.markdown("""
        **Features:**
        ✅ QR code with all part information  
        ✅ Professional layout with borders  
        ✅ Optimized space utilization  
        ✅ One sticker per page  
        ✅ Ready for printing  
        """)

if __name__ == "__main__":
    main()
