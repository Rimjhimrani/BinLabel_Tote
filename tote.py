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

def find_column(df, keywords):
    """Find a column in the DataFrame that matches any of the keywords (case-insensitive)"""
    cols = df.columns.tolist()
    for keyword in keywords:
        for col in cols:
            if isinstance(col, str) and keyword.upper() in col.upper():
                return col
    return None

def extract_line_location_components(row, columns):
    """Extract components for Line Location (L.LOC) from specific columns"""
    components = [''] * 7
    
    # Extract values from respective columns
    if columns['model'] and columns['model'] in row:
        components[0] = str(row[columns['model']]) if pd.notna(row[columns['model']]) else ""
    
    if columns['station_no'] and columns['station_no'] in row:
        components[1] = str(row[columns['station_no']]) if pd.notna(row[columns['station_no']]) else ""
    
    if columns['rack'] and columns['rack'] in row:
        components[2] = str(row[columns['rack']]) if pd.notna(row[columns['rack']]) else ""
    
    # Extract first digit from separate column
    if columns['rack_no_1st'] and columns['rack_no_1st'] in row:
        components[3] = str(row[columns['rack_no_1st']]) if pd.notna(row[columns['rack_no_1st']]) else ""
    
    # Extract second digit from separate column
    if columns['rack_no_2nd'] and columns['rack_no_2nd'] in row:
        components[4] = str(row[columns['rack_no_2nd']]) if pd.notna(row[columns['rack_no_2nd']]) else ""
    
    if columns['level'] and columns['level'] in row:
        components[5] = str(row[columns['level']]) if pd.notna(row[columns['level']]) else ""
    
    if columns['cell'] and columns['cell'] in row:
        components[6] = str(row[columns['cell']]) if pd.notna(row[columns['cell']]) else ""
    
    return components

def extract_store_location_components(row, columns):
    """Extract components for Store Location (S.LOC) from ABB columns"""
    components = [''] * 7
    
    # Extract values from ABB columns
    if columns['abb_zone'] and columns['abb_zone'] in row:
        components[0] = str(row[columns['abb_zone']]) if pd.notna(row[columns['abb_zone']]) else ""
    
    if columns['abb_location'] and columns['abb_location'] in row:
        components[1] = str(row[columns['abb_location']]) if pd.notna(row[columns['abb_location']]) else ""
    
    if columns['abb_floor'] and columns['abb_floor'] in row:
        components[2] = str(row[columns['abb_floor']]) if pd.notna(row[columns['abb_floor']]) else ""
    
    if columns['abb_rack_no'] and columns['abb_rack_no'] in row:
        components[3] = str(row[columns['abb_rack_no']]) if pd.notna(row[columns['abb_rack_no']]) else ""
    
    if columns['abb_level_in_rack'] and columns['abb_level_in_rack'] in row:
        components[4] = str(row[columns['abb_level_in_rack']]) if pd.notna(row[columns['abb_level_in_rack']]) else ""
    
    if columns['abb_cell'] and columns['abb_cell'] in row:
        components[5] = str(row[columns['abb_cell']]) if pd.notna(row[columns['abb_cell']]) else ""
    
    if columns['abb_no'] and columns['abb_no'] in row:
        components[6] = str(row[columns['abb_no']]) if pd.notna(row[columns['abb_no']]) else ""
    
    return components

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
        
    # Identify columns (case-insensitive) - Updated for your exact column names
    original_columns = df.columns.tolist()
    
    # Find basic columns
    part_no_col = find_column(df, ['PART NO', 'PARTNO', 'PART', 'PART_NO', 'PART#'])
    desc_col = find_column(df, ['PART DESC', 'DESC', 'DESCRIPTION', 'NAME', 'PRODUCT_NAME'])
    qty_bin_col = find_column(df, ['QTY/BIN', 'QTY_BIN', 'QTYBIN', 'QTY', 'QUANTITY'])
    
    # Find Line Location columns - Updated for your exact column names
    line_location_columns = {
        'model': find_column(df, ['MODEL', 'BUS MODEL', 'BUS_MODEL', 'BUSMODEL', 'BUS']),
        'station_no': find_column(df, ['STATION NO', 'STATION_NO', 'STATIONNO', 'STATION']),
        'rack': find_column(df, ['RACK']),
        'rack_no_1st': find_column(df, ['RACK NO. (1ST DIGIT)', 'RACK NO (1ST DIGIT)', 'RACK_NO_1ST', 'RACK NO 1ST']),
        'rack_no_2nd': find_column(df, ['RACK NO. (2ND DIGIT)', 'RACK NO (2ND DIGIT)', 'RACK_NO_2ND', 'RACK NO 2ND']),
        'level': find_column(df, ['LEVEL']),
        'cell': find_column(df, ['CELL'])
    }
    
    # Find Store Location (ABB) columns - Updated for your exact column names
    store_location_columns = {
        'abb_zone': find_column(df, ['ABB FOR ZONE', 'ABB_FOR_ZONE', 'ABB ZONE', 'ABB_ZONE', 'ABBZONE', 'ZONE']),
        'abb_location': find_column(df, ['ABB FOR LOCATION', 'ABB_FOR_LOCATION', 'ABB LOCATION', 'ABB_LOCATION', 'ABBLOCATION']),
        'abb_floor': find_column(df, ['ABB FOR FLOOR', 'ABB_FOR_FLOOR', 'ABB FLOOR', 'ABB_FLOOR', 'ABBFLOOR', 'FLOOR']),
        'abb_rack_no': find_column(df, ['ABB FOR RACK NO', 'ABB_FOR_RACK_NO', 'ABB RACK NO', 'ABB_RACK_NO', 'ABBRACKNO', 'ABB RACK']),
        'abb_level_in_rack': find_column(df, ['ABB FOR LEVEL IN RACK', 'ABB_FOR_LEVEL_IN_RACK', 'ABB LEVEL IN RACK', 'ABB_LEVEL_IN_RACK', 'ABBLEVELINRACK', 'ABB LEVEL']),
        'abb_cell': find_column(df, ['ABB FOR CELL', 'ABB_FOR_CELL', 'ABB CELL', 'ABB_CELL', 'ABBCELL']),
        'abb_no': find_column(df, ['ABB FOR NO', 'ABB_FOR_NO', 'ABB NO', 'ABB_NO', 'ABBNO', 'ABB NUMBER'])
    }

    if status_container:
        status_container.write("**Using columns:**")
        status_container.write(f"- Part No: {part_no_col}")
        status_container.write(f"- Description: {desc_col}")
        status_container.write(f"- Qty/Bin: {qty_bin_col}")
        status_container.write("**Line Location columns:**")
        for key, col in line_location_columns.items():
            status_container.write(f"- {key}: {col}")
        status_container.write("**Store Location (ABB) columns:**")
        for key, col in store_location_columns.items():
            status_container.write(f"- {key}: {col}")

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

        # Extract basic data
        part_no = str(row[part_no_col]) if part_no_col and part_no_col in row else ""
        desc = str(row[desc_col]) if desc_col and desc_col in row else ""
        qty_bin = str(row[qty_bin_col]) if qty_bin_col and qty_bin_col in row and pd.notna(row[qty_bin_col]) else ""
        
        # Extract Line Location components
        line_location_parts = extract_line_location_components(row, line_location_columns)
        
        # Extract Store Location components
        store_location_parts = extract_store_location_components(row, store_location_columns)

        # Generate QR code with part information
        qr_data = f"Part No: {part_no}\nDescription: {desc}\nQTY/BIN: {qty_bin}\n"
        qr_data += f"Line Location: {' | '.join(line_location_parts)}\n"
        qr_data += f"Store Location: {' | '.join(store_location_parts)}"
        
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

        store_loc_inner_table = Table(
            [store_location_parts],
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
        
        # Create the inner table for line location parts using the same fixed widths
        line_loc_inner_table = Table(
            [line_location_parts],
            colWidths=inner_col_widths,
            rowHeights=[location_row_height]
        )
        
        line_loc_inner_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.0, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 7)
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

    st.info("👈 Please upload an Excel or CSV file to get started")
    
    # Show expected data format
    st.subheader("📋 Reference For Data Format")
    sample_data = {
        'Part No': ['08-DRA-14-02', 'P0012124-07', 'P0012126-07'],
        'Part Desc': ['BELLOW ASSY. WITH RETAINING CLIP', 'GUARD RING (hirkesh)', 'GUARD RING SEAL (hirkesh)'],
        'Bus model': ['3WC', '3WM', '3WS'],
        'Station No': ['CW40RH', 'CW40RH', 'CW40RH'],
        'Rack': ['R', 'R', 'R'],
        'Rack No (1st digit)': [0, 0, 0],
        'Rack No (2nd digit)': [2, 2, 2],
        'Level': ['A', 'A', 'A'],
        'Cell': [1, 2, 3],
        'ABB ZONE': ['HRD', 'HRD', 'HRD'],
        'ABB LOCATION': ['ABF', 'ABF', 'ABF'],
        'ABB FLOOR': [1, 1, 1],
        'ABB RACK NO': [2, 2, 2],
        'ABB LEVEL IN RACK': ['A', 'A', 'A'],
        'ABB CELL': [1, 2, 3],
        'ABB NO': [1001, 1002, 1003]
    }
    
    sample_df = pd.DataFrame(sample_data)
    st.dataframe(sample_df, use_container_width=True)

    st.markdown("---")
    
    # Column mapping information
    st.subheader("ℹ️ Column Mapping")
    
    # Single column layout - only keeping the previous col2 content
    st.markdown("""
    **Store Location (S.LOC) - 7 boxes:**
    1. **ABB Zone** (ABB ZONE, ABB_ZONE, etc.)
    2. **ABB Location** (ABB LOCATION, ABB_LOCATION, etc.)
    3. **ABB Floor** (ABB FLOOR, ABB_FLOOR, etc.)
    4. **ABB Rack No** (ABB RACK NO, ABB_RACK_NO, etc.)
    5. **ABB Level in Rack** (ABB LEVEL IN RACK, etc.)
    6. **ABB Cell** (ABB CELL, ABB_CELL, etc.)
    7. **ABB No** (ABB NO, ABB_NO, etc.)
    """)
    
    st.markdown("""
    **Features:**
    ✅ Automatic column detection  
    ✅ QR code with all information  
    ✅ Professional layout with borders  
    ✅ 7-box layout for locations  
    ✅ Separate rack digit columns  
    ✅ One sticker per page  
    ✅ Support for new Excel format  
    ✅ ABB location mapping  
    """)
    
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
    
    # Main content area - Single column layout
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
                            file_name=f"{uploaded_file.name.split('.')[0]}_tote_labels.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                        
                        st.success("✅ Tote labels generated successfully!")
                    else:
                        st.error("❌ Failed to generate tote labels")
                        
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
    else:
        pass # Info message is already shown at the top

if __name__ == "__main__":
    main()
