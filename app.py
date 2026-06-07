import openai
import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
import zipfile
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os


# -----------------------------
# OpenRouter Client
# -----------------------------

client = openai.OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=st.secrets["OPENROUTER_API_KEY"]
)


# -----------------------------
# Streamlit Page Setup
# -----------------------------

st.set_page_config(page_title="SAP CPI TDD Generator", layout="wide")

st.title("SAP CPI Technical Design Document Generator")

st.write("Generate a TDD manually or by uploading an SAP CPI iFlow package and mapping sheet.")


# -----------------------------
# User Inputs
# -----------------------------

generation_mode = st.radio(
    "Choose TDD generation option",
    [
        "Manual Entry",
        "Generate using iFlow Package",
        "Generate using iFlow Package + Mapping Sheet"
    ]
)

interface_name = st.text_input("Interface Name")
source_system = st.text_input("Source System")
target_system = st.text_input("Target System")

sender_adapter = ""
receiver_adapter = ""
mapping_logic = ""
error_handling = ""
description = ""

if generation_mode == "Manual Entry":
    sender_adapter = st.text_input("Sender Adapter", placeholder="HTTPS, SFTP, SOAP, IDoc, etc.")
    receiver_adapter = st.text_input("Receiver Adapter", placeholder="OData, REST, SOAP, SFTP, etc.")
    mapping_logic = st.text_area("Mapping Logic")
    error_handling = st.text_area("Error Handling")
    description = st.text_area("Integration Flow Description")

iflow_file = None
mapping_file = None
iflow_summary = ""
mapping_summary = ""

if generation_mode in ["Generate using iFlow Package", "Generate using iFlow Package + Mapping Sheet"]:
    iflow_file = st.file_uploader("Upload SAP CPI iFlow Package ZIP", type=["zip"])

if generation_mode == "Generate using iFlow Package + Mapping Sheet":
    mapping_file = st.file_uploader("Upload Mapping Sheet", type=["xlsx", "xls", "csv"])


# -----------------------------
# Extract iFlow Package Details
# -----------------------------

def extract_iflow_details(uploaded_zip):
    details = []

    try:
        zip_bytes = BytesIO(uploaded_zip.read())

        with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
            file_names = zip_ref.namelist()

            details.append("Files found in iFlow package:")

            for name in file_names:
                details.append(f"- {name}")

            for name in file_names:
                lower_name = name.lower()

                if lower_name.endswith(".iflw") or lower_name.endswith(".xml"):
                    try:
                        content = zip_ref.read(name).decode("utf-8", errors="ignore")
                        details.append(f"\nContent extracted from {name}:")
                        details.append(content[:8000])
                    except Exception:
                        details.append(f"Unable to read {name}")

                if lower_name.endswith(".groovy"):
                    try:
                        content = zip_ref.read(name).decode("utf-8", errors="ignore")
                        details.append(f"\nGroovy script found: {name}")
                        details.append(content[:4000])
                    except Exception:
                        details.append(f"Unable to read {name}")

                if lower_name.endswith(".xsl") or lower_name.endswith(".xslt"):
                    try:
                        content = zip_ref.read(name).decode("utf-8", errors="ignore")
                        details.append(f"\nXSLT mapping found: {name}")
                        details.append(content[:4000])
                    except Exception:
                        details.append(f"Unable to read {name}")

                if lower_name.endswith(".wsdl") or lower_name.endswith(".xsd"):
                    try:
                        content = zip_ref.read(name).decode("utf-8", errors="ignore")
                        details.append(f"\nSchema/WSDL found: {name}")
                        details.append(content[:4000])
                    except Exception:
                        details.append(f"Unable to read {name}")

                if lower_name.endswith(".properties") or lower_name.endswith(".prop"):
                    try:
                        content = zip_ref.read(name).decode("utf-8", errors="ignore")
                        details.append(f"\nProperties file found: {name}")
                        details.append(content[:4000])
                    except Exception:
                        details.append(f"Unable to read {name}")

        return "\n".join(details)

    except Exception as e:
        return f"Error reading iFlow package: {e}"


# -----------------------------
# Extract Mapping Sheet
# -----------------------------

def extract_mapping_sheet(uploaded_file):
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        preview = df.head(50).to_string(index=False)

        return df, preview

    except Exception as e:
        return None, f"Error reading mapping sheet: {e}"


# -----------------------------
# Create Architecture Diagram
# -----------------------------

def create_architecture_diagram(
    interface_name,
    source_system,
    target_system,
    sender_adapter,
    receiver_adapter
):
    width = 1500
    height = 550

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)

    try:
        font = ImageFont.truetype("arial.ttf", 22)
        small_font = ImageFont.truetype("arial.ttf", 18)
        title_font = ImageFont.truetype("arial.ttf", 32)
    except Exception:
        font = ImageFont.load_default()
        small_font = ImageFont.load_default()
        title_font = ImageFont.load_default()

    draw.text((500, 30), "SAP CPI Architecture Diagram", fill="black", font=title_font)

    boxes = [
        ("Source System\n" + (source_system or "To be confirmed"), 50, 190, 190, 110),
        ("Sender Adapter\n" + (sender_adapter or "To be confirmed"), 280, 190, 190, 110),
        ("SAP CPI /\nIntegration Suite\n" + (interface_name or "Interface"), 530, 170, 230, 150),
        ("Mapping /\nTransformation", 820, 190, 190, 110),
        ("Receiver Adapter\n" + (receiver_adapter or "To be confirmed"), 1050, 190, 190, 110),
        ("Target System\n" + (target_system or "To be confirmed"), 1280, 190, 190, 110),
    ]

    for text, x, y, box_width, box_height in boxes:
        draw.rectangle((x, y, x + box_width, y + box_height), outline="black", width=3)
        draw.multiline_text((x + 12, y + 25), text, fill="black", font=small_font, spacing=6)

    arrows = [
        (240, 245, 280, 245),
        (470, 245, 530, 245),
        (760, 245, 820, 245),
        (1010, 245, 1050, 245),
        (1240, 245, 1280, 245),
    ]

    for x1, y1, x2, y2 in arrows:
        draw.line((x1, y1, x2, y2), fill="black", width=3)
        draw.polygon(
            [(x2, y2), (x2 - 12, y2 - 8), (x2 - 12, y2 + 8)],
            fill="black"
        )

    # Monitoring box
    draw.rectangle((580, 390, 950, 470), outline="black", width=3)
    draw.multiline_text(
        (610, 415),
        "Monitoring, Logging\nand Error Handling",
        fill="black",
        font=small_font,
        spacing=6
    )

    # Arrow from CPI to monitoring
    draw.line((650, 320, 700, 390), fill="black", width=3)
    draw.polygon(
        [(700, 390), (688, 382), (710, 378)],
        fill="black"
    )

    diagram_file = "architecture_diagram.png"
    image.save(diagram_file)

    return diagram_file


# -----------------------------
# Call OpenRouter AI
# -----------------------------

def call_ai(prompt):
    try:
        response = client.chat.completions.create(
            model="openai/gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an SAP CPI Integration Architect. "
                        "Generate accurate technical design documents. "
                        "Do not invent missing information."
                    )
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.2
        )

        return response.choices[0].message.content

    except Exception as e:
        return f"Error calling OpenRouter: {e}"


# -----------------------------
# Create Word Document
# -----------------------------

def create_word_document(content, mapping_df=None):
    doc = Document()

    doc.add_heading("SAP CPI Technical Design Document", 0)

    doc.add_heading("1. Interface Details", level=1)
    doc.add_paragraph(f"Interface Name: {interface_name}")
    doc.add_paragraph(f"Source System: {source_system}")
    doc.add_paragraph(f"Target System: {target_system}")
    doc.add_paragraph(f"Sender Adapter: {sender_adapter or 'To be confirmed'}")
    doc.add_paragraph(f"Receiver Adapter: {receiver_adapter or 'To be confirmed'}")

    doc.add_heading("2. Architecture Diagram", level=1)

    try:
        diagram_file = create_architecture_diagram(
            interface_name,
            source_system,
            target_system,
            sender_adapter,
            receiver_adapter
        )
        doc.add_picture(diagram_file, width=Inches(6.5))
    except Exception as e:
        doc.add_paragraph(f"Architecture diagram could not be generated: {e}")

    doc.add_page_break()

    doc.add_heading("3. Generated Technical Design Document", level=1)
    doc.add_paragraph(content)

    if mapping_df is not None:
        doc.add_page_break()
        doc.add_heading("Appendix A - Mapping Sheet", level=1)

        table = doc.add_table(rows=1, cols=len(mapping_df.columns))
        table.style = "Table Grid"

        header_cells = table.rows[0].cells

        for i, column in enumerate(mapping_df.columns):
            header_cells[i].text = str(column)

        for _, row in mapping_df.head(100).iterrows():
            row_cells = table.add_row().cells

            for i, value in enumerate(row):
                row_cells[i].text = "" if pd.isna(value) else str(value)

        if len(mapping_df) > 100:
            doc.add_paragraph(
                "Note: Only first 100 rows of mapping sheet are included in this document."
            )

    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    return file_stream


# -----------------------------
# Generate TDD Button
# -----------------------------

if st.button("Generate TDD"):
    mapping_df = None

    if not interface_name or not source_system or not target_system:
        st.warning("Please enter Interface Name, Source System, and Target System.")

    else:
        if iflow_file is not None:
            with st.spinner("Reading iFlow package..."):
                iflow_summary = extract_iflow_details(iflow_file)

        if mapping_file is not None:
            with st.spinner("Reading mapping sheet..."):
                mapping_df, mapping_summary = extract_mapping_sheet(mapping_file)

        prompt = f"""
You are an SAP CPI Integration Architect.

Generate a detailed Technical Design Document for the following SAP CPI interface.

Interface Name: {interface_name}
Source System: {source_system}
Target System: {target_system}

Manual Details:
Sender Adapter: {sender_adapter}
Receiver Adapter: {receiver_adapter}
Mapping Logic: {mapping_logic}
Error Handling: {error_handling}
Integration Flow Description: {description}

Extracted iFlow Package Details:
{iflow_summary}

Mapping Sheet Details:
{mapping_summary}

Create the document with these sections:

1. Document Purpose
2. Interface Overview
3. Business Requirement Summary
4. Architecture Diagram Description
5. Source System Details
6. Target System Details
7. SAP CPI iFlow Design
8. Sender Adapter Configuration
9. Receiver Adapter Configuration
10. Message Mapping and Transformation Logic
11. Routing Logic
12. Exception Handling
13. Retry and Reprocessing Strategy
14. Security and Authentication
15. Monitoring and Logging
16. Externalized Parameters
17. Mapping Sheet Summary
18. Assumptions
19. Dependencies
20. Testing Scope
21. Test Scenarios
22. Deployment Checklist
23. Appendix Reference

Important rules:
- Use the iFlow package details where available.
- Use the mapping sheet details where available.
- Only use information provided by the user, iFlow package, or mapping sheet.
- Do not invent missing information.
- If any information is missing, mention "To be confirmed".
- Do not include passwords, secrets, API keys, or credentials.
"""

        with st.spinner("Generating Technical Design Document..."):
            tdd_content = call_ai(prompt)

        st.subheader("Generated TDD")
        st.write(tdd_content)

        # Generate and show architecture diagram in Streamlit
        try:
            diagram_file = create_architecture_diagram(
                interface_name,
                source_system,
                target_system,
                sender_adapter,
                receiver_adapter
            )

            st.subheader("Architecture Diagram")
            st.image(diagram_file)

            with open(diagram_file, "rb") as img:
                st.download_button(
                    label="Download Architecture Diagram",
                    data=img,
                    file_name="architecture_diagram.png",
                    mime="image/png"
                )

        except Exception as e:
            st.error(f"Architecture diagram could not be generated: {e}")

        # Optional review/edit before download
        edited_tdd = st.text_area(
            "Review/Edit Generated TDD before download",
            value=tdd_content,
            height=500
        )

        word_file = create_word_document(edited_tdd, mapping_df)

        st.download_button(
            label="Download TDD as Word Document",
            data=word_file,
            file_name=f"{interface_name}_TDD.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )