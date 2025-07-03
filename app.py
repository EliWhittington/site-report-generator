import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from PIL import Image, ImageOps
from datetime import date
import os
import io
import re
import tempfile

# === Streamlit App Config ===
st.set_page_config(page_title="Site Report Generator", layout="wide")
st.title("📷 Site Observation Report Generator")

# === Helper: Extract number from filename for sorting ===
def extract_number(filename):
    match = re.search(r'(\d+)', filename)
    return int(match.group(1)) if match else float('inf')

# === Resize image ===
def resize_image(image_bytes, max_dim):
    img = Image.open(io.BytesIO(image_bytes))
    img = ImageOps.exif_transpose(img)
    img = img.convert("RGB")
    img.thumbnail((max_dim, max_dim))
    return img

# === Generate the report ===
def generate_report(images, weather, subcontractors, areas, max_dim, quality):
    doc = Document()

    # === Title ===
    title_paragraph = doc.add_paragraph()
    title_paragraph.alignment = 1  # Center
    run = title_paragraph.add_run("Project Observation Report")
    run.font.name = 'Calibri'
    run.font.size = Pt(28)
    doc.add_paragraph()

    today_str = date.today().strftime("%B %d, %Y")

    # === Static Fields ===
    project_name = "National Fire Cache Building                                        -"
    user_name = "Eli Whittington12"
    email = "eli.whittington@pc.gc.ca"
    address = "200 Hawk Ave, Banff AB T1L1K2"
    review_date = today_str

    def add_field(label, value):
        p = doc.add_paragraph()
        run_label = p.add_run(f"{label}: ")
        run_label.bold = True
        run_label.font.name = 'Calibri'
        run_label.font.size = Pt(12)
        run_value = p.add_run(value)
        run_value.font.name = 'Calibri'
        run_value.font.size = Pt(12)

    add_field("Project", project_name)
    add_field("Name", user_name)
    add_field("Email", email)
    add_field("Review Date", review_date)
    add_field("Report Date", today_str)
    add_field("Address", address)
    add_field("Conditions", f"{weather} °C")

    doc.add_paragraph()  # Spacing

    # === Subcontractors ===
    doc.add_paragraph("Subcontractors Present:", style="Heading 2")
    for s in subcontractors:
        doc.add_paragraph(f"    - {s}")

    doc.add_paragraph("Areas of Work/Progress:", style="Heading 2")
    for a in areas:
        doc.add_paragraph(f"    - {a}")

    doc.add_page_break()

    # === Insert 2 Images Per Page ===
    # === Insert 2 Images Per Page (no forced page breaks) ===
    for i in range(0, len(images), 2):
        # First image
        doc.add_picture(images[i], width=Inches(5.93), height=Inches(4.45))
        doc.paragraphs[-1].alignment = 1  # Center

        # Second image, if available
        if i + 1 < len(images):
            doc.add_picture(images[i + 1], width=Inches(5.93), height=Inches(4.45))
            doc.paragraphs[-1].alignment = 1  # Center

        # Small vertical space between image pairs
        #doc.add_paragraph()



    # === Save to BytesIO for download ===
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# === Streamlit Inputs ===
weather = st.text_input("Weather Conditions (°C)")
max_dim = st.number_input("Max image dimension (px)", min_value=500, max_value=5000, value=1300)
quality = st.slider("JPEG Quality (1 = small file, 100 = high quality)", 1, 100, 95)

subcontractors = st.text_area("Subcontractors (one per line)").split("\n")
areas = st.text_area("Areas of Work (one per line)").split("\n")

uploaded_files = st.file_uploader("Upload Site Images", type=['jpg', 'jpeg'], accept_multiple_files=True)

# === Button to Generate Report ===
if st.button("Generate Report"):
    if not uploaded_files:
        st.warning("Please upload at least one image.")
    else:
        with st.spinner("Generating report..."):
            # Sort images by number in filename
            uploaded_files.sort(key=lambda f: extract_number(f.name))

            temp_images = []
            for file in uploaded_files:
                img = resize_image(file.read(), max_dim)
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
                img.save(temp_file.name, format="JPEG", quality=quality, optimize=True)
                temp_images.append(temp_file.name)

            report_bytes = generate_report(temp_images, weather, subcontractors, areas, max_dim, quality)
            st.success("Report generated!")
            st.download_button("📄 Download Report", report_bytes, file_name="Progress_Report.docx")
