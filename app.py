import streamlit as st
import requests
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Key Mapping for readability
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
        "Part A - Valuation of Land": f"https://valuerkkda.in/wp-json/part-a/part-a/?_post_id={post_id}",
        "Part C - Extra Items": f"https://valuerkkda.in/wp-json/part-c/part-c/?_post_id={post_id}",
        "Part D - Amenities": f"https://valuerkkda.in/wp-json/part-d/part-d/?_post_id={post_id}",
        "Part E - Miscellaneous": f"https://valuerkkda.in/wp-json/part-e/part-e/?_post_id={post_id}",
        "Part F - Services": f"https://valuerkkda.in/wp-json/get_releted_Part_F/get-releted-part-f-/?_post_id={post_id}",
        "Specifications of Construction": f"https://valuerkkda.in/wp-json/specifications/part-specifications/?_post_id={post_id}",
    }

    document = Document()
    document.add_heading(f"Valuation Report for Post ID: {post_id}", level=1)

    sections = {
        "Part A - Valuation of Land": api_to_dataframe(api_urls["Part A - Valuation of Land"]),
        "Specifications of Construction": split_and_format_specifications(api_to_dataframe(api_urls["Specifications of Construction"])),
        "Part C - Extra Items": api_to_dataframe(api_urls["Part C - Extra Items"]),
        "Part D - Amenities": api_to_dataframe(api_urls["Part D - Amenities"]),
        "Part E - Miscellaneous": api_to_dataframe(api_urls["Part E - Miscellaneous"]),
        
    }

    for section, df in sections.items():
        if not df.empty:
            st.subheader(section)
            st.table(df)
            
            document.add_heading(section, level=2)
            docx_table = document.add_table(rows=1, cols=len(df.columns))
            
            # Apply table style
            apply_table_style(docx_table)

            hdr_cells = docx_table.rows[0].cells
            for i, col_name in enumerate(df.columns):
                hdr_cells[i].text = col_name
                hdr_cells[i].paragraphs[0].runs[0].bold = True

            for _, row in df.iterrows():
                row_cells = docx_table.add_row().cells
                for i, value in enumerate(row):
                    row_cells[i].text = str(value)
                    row_cells[i].paragraphs[0].runs[0].font.size = Pt(10)

    docx_buffer = BytesIO()
    document.save(docx_buffer)
    docx_buffer.seek(0)

    st.download_button("Download Report", data=docx_buffer, file_name=f"valuation_report_{post_id}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
