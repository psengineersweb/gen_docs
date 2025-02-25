import streamlit as st
import requests
import pandas as pd
from io import BytesIO
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement, ns
post_id=0
# Key Mapping for readability


def apply_header_footer(doc):
    """Applies the same header and footer to all sections of the document."""
    for section in doc.sections:
        section.start_type = 2  # Ensure new sections start on a new page

        # === HEADER ===
        header = section.header
        header.is_linked_to_previous = False
        
        # Create a table with a specified width inside the header
        header_table = header.add_table(rows=1, cols=2, width=Inches(6.5))
        header_table.allow_autofit = False

        # Set explicit column widths
        cell_text = header_table.cell(0, 0)
        cell_logo = header_table.cell(0, 1)
        cell_text.width = Inches(5)  # 75% width
        cell_logo.width = Inches(1.5)  # 25% width

        # HEADER TEXT
        title = cell_text.paragraphs[0]
        title.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        run = title.add_run("KOUSHIK KUMAR DAS")
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(0, 0, 255)  # Blue text

        details = cell_text.add_paragraph()
        details.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        text = (
            "M.SC. REAL ESTATE VALUATION, M.I.S.(VALUATION-SURVEYING), AMISE (CIVIL ENGG), "
            "LIFE ASSOCIATE OF THE INSTITUTION OF SURVEYORS (VALUATION - SURVEYING), F.I.V., "
            "APPROVED VALUER, INSTITUTION OF VALUERS, BANKS, INSURANCE, FINANCIAL AND INDUSTRIAL "
            "CORPORATIONS & APPROVED VALUER (CAT-1) OF IMMOVABLE PROPERTY & ENGINEER COMMISSIONER’ "
            "OF THE OFFICE OF THE CITY CIVIL COURT, CALCUTTA."
        )
        run = details.add_run(text)
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0, 0, 255)

        # HEADER LOGO
        paragraph_logo = cell_logo.paragraphs[0]
        paragraph_logo.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        run_logo = paragraph_logo.add_run()

        try:
            run_logo.add_picture("logo.png", width=Inches(1.5))
        except FileNotFoundError:
            print("Warning: Logo file not found. Add 'logo.png' to the working directory.")

        # === FOOTER ===
        footer = section.footer
        footer.is_linked_to_previous = False
        footer_paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Center align the footer

        footer_text = (
            "OFFICE: 20/1, CHETLA HAT ROAD, P.O. ALIPORE (H.O.), KOLKATA – 700027\n"
            "MOBILE NOS: 9903068614, 9830776121, 9433240397,\n"
            "E-MAIL- kkdareport@gmail.com valuerkoushik@yahoo.co.in"
        )

        run_footer = footer_paragraph.add_run(footer_text)
        run_footer.italic = True
        run_footer.font.size = Pt(10)
        run_footer.font.color.rgb = RGBColor(0, 0, 255)  # Blue color


key_map = {
    # Part A - Valuation of Land
    "size_of_plot": "Size of Plot",
    "north": "North",
    "south": "South",
    "east": "East",
    "west": "West",
    "total_extent_of_the_plot": "Total Extent of the Plot",
    "prevailing_market_rate": "Prevailing Market Rate",
    "guideline_rate_obtained_from_the_registrars_office": "Guideline Rate",
    "assessedadopted_rate_of_valuation": "Assessed/Adopted Rate",
    "estimated_value_of_land": "Estimated Value of Land",

    # Part C - Extra Items
    "portico": "Portico",
    "ornamental_front_door": "Ornamental Front Door",
    "sit_out_verandah_with_steel_grills": "Sit Out/Verandah with Steel Grills",
    "overhead_water_tank": "Overhead Water Tank",
    "extra_steel_collapsible_gates": "Extra Steel/Collapsible Gates",

    # Part D - Amenities
    "wardrobes": "Wardrobes",
    "glazed_tiles": "Glazed Tiles",
    "extra_sinks_and_bath_tub": "Extra Sinks and Bath Tub",
    "marble__ceramic_tiles_flooring": "Interior Decorations",
    "interior_decorations": "Interior Decorations",
    "architectural_elevation_works": "Architectural Elevation Works",
    "panelling_works": "Panelling Works",
    "aluminium_works": "Aluminium Works",
    "aluminium_hand_rails": "Aluminium Hand Rails",
    "false_ceiling": "False Ceiling",

    # Part E - Miscellaneous
    "separate_toilet_room": "Separate Toilet Room",
    "separate_lumber_room": "Separate Lumber Room",
    "separate_water_tank_sump": "Separate Water Tank/Sump",
    "trees_gardening": "Trees, Gardening",

    # Part F - Services
    "water_supply_arrangements": "Water Supply Arrangements",
    "drainage_arrangements": "Drainage Arrangements",
    "compound_wall": "Compound Wall",
    "c_b_deposits_fittings_etc": "C. B. Deposits, Fittings etc.",
    "pavement": "Pavement"
}

# Function to apply table style with borders
from docx.oxml import OxmlElement, ns

def apply_table_style(table):
    """Apply borders to a table in a Word document."""
    tbl = table._element

    # Correctly find `w:tblPr` using the namespace
    tblPr = tbl.find(ns.qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    # Remove existing `w:tblBorders` if present
    tblBorders = tblPr.find(ns.qn("w:tblBorders"))
    if tblBorders is not None:
        tblPr.remove(tblBorders)

    # Create new `w:tblBorders`
    tblBorders = OxmlElement("w:tblBorders")

    # Define borders for each side
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(ns.qn("w:val"), "single")  # Solid border
        border.set(ns.qn("w:sz"), "8")  # Thickness
        border.set(ns.qn("w:space"), "0")  # No spacing
        border.set(ns.qn("w:color"), "000000")  # Black color
        tblBorders.append(border)

    # Add the borders to the table properties
    tblPr.append(tblBorders)



def api_to_dataframe(api_url):
    excluded_keys = {
        "ID", "post_author", "post_date", "post_date_gmt", "post_content", "post_title",
        "post_excerpt", "post_status", "comment_status", "ping_status", "post_password",
        "post_name", "to_ping", "pinged", "post_modified", "post_modified_gmt",
        "post_content_filtered", "post_parent", "guid", "menu_order", "post_type",
        "post_mime_type", "comment_count", "filter", "meta_ID", "object_ID", "_ID", 
        "created", "rel_id", "parent_rel", "parent_object_id", "child_object_id",
    }   
    try:
        response = requests.get(api_url)
        response.raise_for_status()
        data = response.json()
        
        if not data:
            st.warning(f"No data found for {api_url}")
            return pd.DataFrame(columns=['Key', 'Value'])

        if isinstance(data, dict):
            filtered_data = {key_map.get(k, k): v for k, v in data.items() if k not in excluded_keys}
            df = pd.DataFrame(filtered_data.items(), columns=['Key', 'Value'])
        elif isinstance(data, list):
            flattened_data = []
            for item in data:
                filtered_item = {key_map.get(k, k): v for k, v in item.items() if k not in excluded_keys}
                flattened_data.extend(filtered_item.items())
            df = pd.DataFrame(flattened_data, columns=['Key', 'Value'])
        else:
            raise ValueError("Unsupported JSON structure")

        return df
    except requests.RequestException as e:
        st.error(f"API request failed: {e}")
    except ValueError as ve:
        st.error(f"Data processing error: {ve}")
    
    return pd.DataFrame(columns=['Key', 'Value'])

def split_and_format_specifications(specifications_df):
    if specifications_df.empty:
        return pd.DataFrame(columns=['Description', 'Ground Floor', 'Others Floor'])

    def split_key(key):
        parts = key.rsplit('_', 2)
        if len(parts) >= 2:
            desc, floor = parts[0], parts[1]
            return desc.replace('_', ' ').capitalize(), floor.replace('_', ' ').capitalize()
        return key.replace('_', ' ').capitalize(), ""

    specifications_df[['Description', 'Floor']] = specifications_df['Key'].apply(lambda x: pd.Series(split_key(x)))
    
    pivot_df = specifications_df.drop(columns=['Key']).pivot_table(
        index='Description', 
        columns='Floor', 
        values='Value', 
        aggfunc='first'
    ).reset_index()
    
    pivot_df.columns.name = None
    pivot_df = pivot_df.rename(columns={'Ground': 'Ground Floor', 'Others': 'Others Floor'})

    return pivot_df

st.title("Valuation Report Generator")
post_id = st.text_input("Enter Post ID:")

if post_id:
    api_urls = {
        
        "Generel": f"https://valuerkkda.in/wp-json/generel/generel/?_post_id={post_id}",
        "Part A - Valuation of Land": f"https://valuerkkda.in/wp-json/part-a/part-a/?_post_id={post_id}",
        "Part - B (Valuation of Building)": f"https://valuerkkda.in/wp-json/part-b/part-b/?_post_id={post_id}",
        "Part C - Extra Items": f"https://valuerkkda.in/wp-json/part-c/part-c/?_post_id={post_id}",
        "Part D - Amenities": f"https://valuerkkda.in/wp-json/part-d/part-d/?_post_id={post_id}",
        "Part E - Miscellaneous": f"https://valuerkkda.in/wp-json/part-e/part-e/?_post_id={post_id}",
        "Part F - Services": f"https://valuerkkda.in/wp-json/get_releted_Part_F/get-releted-part-f-/?_post_id={post_id}",
        "Specifications of Construction": f"https://valuerkkda.in/wp-json/specifications/part-specifications/?_post_id={post_id}",
    }

    document = Document()

    apply_header_footer(document)

    sections = {
        "VALUATION REPORT (IN RESPECT OF LAND)": "VALUATION REPORT (IN RESPECT OF LAND)",
        "Owners": "Owners",
        "General": api_to_dataframe(api_urls["Generel"]),
        "Part A - Valuation of Land": api_to_dataframe(api_urls["Part A - Valuation of Land"]),
        "Part - B (Valuation of Building)": api_to_dataframe(api_urls["Part - B (Valuation of Building)"]),
        "Specifications of Construction": split_and_format_specifications(api_to_dataframe(api_urls["Specifications of Construction"])),
        "Details of Valuation": "Details of Valuation",
        "Part C - Extra Items": api_to_dataframe(api_urls["Part C - Extra Items"]),
        "Part D - Amenities": api_to_dataframe(api_urls["Part D - Amenities"]),
        "Part E - Miscellaneous": api_to_dataframe(api_urls["Part E - Miscellaneous"]),
        "Part F - Services": api_to_dataframe(api_urls["Part F - Services"]),
        "PRESENT VALUE OF SAID PROPERTY": "PRESENT VALUE OF SAID PROPERTY",
        "CERTIFICATE OF STABILITY": "CERTIFICATE OF STABILITY",
        "VETTED ESTIMATE": "VETTED ESTIMATE",
        "Format of undertaking to be submitted by Individuals/ proprietor/ partners/ directors DECLARATION- CUM- UNDERTAKING": "Format of undertaking to be submitted by Individuals/ proprietor/ partners/ directors DECLARATION- CUM- UNDERTAKING",
        "Further, I hereby provide the following information.": "Further, I hereby provide the following information.",
        "MODEL CODE OF CONDUCT FOR VALUERS": "MODEL CODE OF CONDUCT FOR VALUERS",
    }

    

    for section, content in sections.items():
        st.subheader(section)
        document.add_heading(section, level=2)

        # Check if content is a DataFrame
        if isinstance(content, pd.DataFrame):
            if not content.empty:
                st.table(content)

                # Create table in docx
                docx_table = document.add_table(rows=1, cols=len(content.columns))

                # Apply table style if function exists
                if "apply_table_style" in globals():
                    apply_table_style(docx_table)

                hdr_cells = docx_table.rows[0].cells
                for i, col_name in enumerate(content.columns):
                    hdr_cells[i].text = col_name
                    hdr_cells[i].paragraphs[0].runs[0].bold = True

                for _, row in content.iterrows():
                    row_cells = docx_table.add_row().cells
                    for i, value in enumerate(row):
                        row_cells[i].text = str(value)
                        row_cells[i].paragraphs[0].runs[0].font.size = Pt(10)
        else:
            st.write(content)
            document.add_paragraph(content)

    # Save the Word document
    docx_buffer = BytesIO()
    document.save(docx_buffer)
    docx_buffer.seek(0)
try:
    st.download_button("Download Report", data=docx_buffer, file_name=f"valuation_report_{post_id}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
except: pass