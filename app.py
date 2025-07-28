import streamlit as st
import pandas as pd
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom import minidom
import io  # Added for in-memory file handling

# --- Corporate Theme Settings ---
st.set_page_config(page_title="Org Chart Generator", layout="wide")

# --- Corporate Header with local image ---
col1, col2 = st.columns([10, 1])
with col1:
    st.markdown("<h1 style='color: #8B1C3F; font-size: 26px; font-weight: 700;'>Organizational Chart Generator</h1>", unsafe_allow_html=True)
with col2:
    st.image("icon.jpg", width=48)

# --- Page Styling ---
st.markdown("""
    <style>
    body {
        background-color: #f4f4f4;
        font-family: 'Segoe UI', sans-serif;
    }
    .main {
        background-color: #ffffff;
        padding: 30px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border: 1px solid #ddd;
    }
    h1 {
        color: #8B1C3F;
        font-size: 28px;
        font-weight: 700;
        border-bottom: 2px solid #F5C014;
        padding-bottom: 12px;
        margin-bottom: 24px;
    }
    .stButton button {
        background-color: #8B1C3F;
        color: white;
        font-weight: 600;
        border-radius: 5px;
        border: none;
        padding: 10px 20px;
        font-size: 14px;
    }
    .stButton button:hover {
        background-color: #6e0f2c;
    }
    .stDownloadButton button {
        background-color: #F5C014;
        color: #000;
        font-weight: 600;
        border-radius: 5px;
        padding: 10px 20px;
        font-size: 14px;
        border: none;
    }
    .stDownloadButton button:hover {
        background-color: #e0b000;
    }
    .stFileUploader label, .stSelectbox label, .stExpander label {
        color: #333;
        font-weight: 600;
        font-size: 14px;
    }
    .stMarkdown h3 {
        color: #8B1C3F;
        font-size: 18px;
        font-weight: 600;
    }
    .stExpander {
        border: 1px solid #ddd;
        border-radius: 5px;
        background-color: #fcfcfc;
    }
    .footer {
        color: #777;
        font-size: 13px;
        margin-top: 30px;
    }
    hr {
        border-top: 1px solid #ccc;
        margin: 30px 0;
    }
    </style>
""", unsafe_allow_html=True)

# --- Instructions ---
with st.expander("📘 How to Use the Tool", expanded=False):
    st.markdown("""
    ✅ **1. Prepare Your Excel** with the following columns:
    - `Title` — Employee title
    - `Manager Title` — Direct supervisor
    - Location columns (e.g., `Riyadh`, `Dammam`) with `1` or `0`

    ✅ **2. Upload the Excel File** below.

    ✅ **3. Assign Roots (if needed):** For employees without managers, choose a root.

    ✅ **4. Download your .drawio File:** and open it at [diagrams.net](https://app.diagrams.net)

    ✅ **5. Format on Draw.io:**
    - Arrange > Layout > Org Chart / Vertical Tree
    - Style edges: straight
    - Customize boxes and fonts if needed.
    """)

# --- Node Color Levels ---
LEVEL_COLORS = {
    0: "#a3c9e2", 1: "#b1d8b7", 2: "#f8cfa1", 3: "#e6b8b7", 4: "#d0d9f2",
    5: "#e2ded0", 6: "#f2d7e2", 7: "#d4eac8", 8: "#fbe8c9", 9: "#ccd9f9", 10: "#e9e6df"
}

# --- Org Chart Builder ---
def build_drawio_xml(df, location, root_overrides):
    df = df[df[location] > 0].copy()
    mxfile = Element('mxfile', host='app.diagrams.net')
    diagram = SubElement(mxfile, 'diagram', name=f'OrgChart_{location}')
    mxGraphModel = SubElement(diagram, 'mxGraphModel')
    root = SubElement(mxGraphModel, 'root')
    SubElement(root, 'mxCell', id="0")
    SubElement(root, 'mxCell', id="1", parent="0")

    id_map = {}
    counter = 2
    height_gap = 110
    width_gap = 200

    def get_level(title):
        level = 0
        current = title
        while pd.notna(df[df["Title"] == current]["Manager Title"].values[0]):
            current = df[df["Title"] == current]["Manager Title"].values[0]
            if current not in df["Title"].values:
                break
            level += 1
        return level

    for title, new_manager in root_overrides.items():
        df.loc[df["Title"] == title, "Manager Title"] = new_manager if new_manager != "None" else None

    levels = {}
    for _, row in df.iterrows():
        level = get_level(row["Title"])
        levels.setdefault(level, []).append(row["Title"])

    for level in sorted(levels.keys()):
        titles = levels[level]
        for i, title in enumerate(titles):
            x = width_gap * i
            y = height_gap * level
            node_id = str(counter)
            id_map[title] = node_id
            fill = LEVEL_COLORS.get(level, "#ffffff")
            cell = SubElement(root, 'mxCell',
                              id=node_id,
                              value=f"{title}",
                              style=f"rounded=1;whiteSpace=wrap;html=1;fillColor={fill};strokeColor=#444;fontSize=12;",
                              vertex="1",
                              parent="1")
            geometry = SubElement(cell, 'mxGeometry')
            geometry.set('x', str(x))
            geometry.set('y', str(y))
            geometry.set('width', "160")
            geometry.set('height', "60")
            geometry.set('as', "geometry")
            counter += 1

    for _, row in df.iterrows():
        manager = row["Manager Title"]
        employee = row["Title"]
        if pd.notna(manager) and manager in id_map and employee in id_map:
            edge = SubElement(root, 'mxCell',
                              id=str(counter),
                              style="endArrow=none;strokeColor=#888;",
                              edge="1",
                              parent="1",
                              source=id_map[manager],
                              target=id_map[employee])
            geometry = SubElement(edge, 'mxGeometry')
            geometry.set('relative', "1")
            geometry.set('as', "geometry")
            counter += 1

    return minidom.parseString(tostring(mxfile)).toprettyxml(indent="  ")

# --- Upload Excel ---
st.markdown("### 📤 Upload Excel File")
uploaded_file = st.file_uploader("Upload your Excel sheet here", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if not {"Title", "Manager Title"}.issubset(df.columns):
        st.error("❌ Excel must contain 'Title' and 'Manager Title' columns.")
    else:
        location_columns = df.columns[2:]
        st.success("✅ Detected Locations: " + ", ".join(location_columns))

        for location in location_columns:
            st.markdown(f"#### 📍 Location: **{location}**")
            location_df = df[df[location] > 0].copy()
            orphaned_titles = location_df[
                ~location_df["Manager Title"].isin(location_df["Title"]) & pd.notna(location_df["Manager Title"])
            ]["Title"].tolist()

            root_overrides = {}
            if orphaned_titles:
                st.warning(f"⚠️ Missing managers for: {', '.join(orphaned_titles)}")
                for title in orphaned_titles:
                    options = ["None"] + sorted(location_df["Title"].tolist())
                    root_choice = st.selectbox(f"🔁 Who should '{title}' report to?", options, key=f"{location}_{title}")
                    root_overrides[title] = root_choice

            # Generate XML string
            xml_str = build_drawio_xml(df, location, root_overrides)
            
            # Use BytesIO for in-memory file handling
            xml_bytes = io.BytesIO(xml_str.encode('utf-8'))
            
            # Provide download button with in-memory file
            st.download_button(
                label=f"📥 Download Org Chart for {location}",
                data=xml_bytes,
                file_name=f"org_chart_{location}.drawio",
                mime="application/xml"
            )

# --- Footer ---
st.markdown("""
<br><hr><center style='color:#aaa; font-size: 14px;'>© Zahra Aljanabi — All Rights Reserved</center>
""", unsafe_allow_html=True)
