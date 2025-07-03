import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from PIL import Image, ImageOps
from datetime import date
import os
import io
import re
import tempfile

st.set_page_config(page_title="Test App", layout="wide")
st.title("✅ The app is running!")

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

# === Build the report ===
def generate_report(images, weather, subcontractors, areas, max_dim, quality):
    doc = Document()
    # Title
    title_paragraph = doc.add_paragraph()
    title_paragraph.alignment = 1  # Center
    run = title_paragraph.add_run("Project Observation Report")
    run.font.name = 'Calibri'
    run.font.size = Pt(28)
    doc.add_paragraph()

    today_str = date.today().strftime("%B %d, %Y")

    def add_field(label, value):
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(value)

    add_field("Report Date", today_str)
    add_field("Conditions", f"{weather} °C")

    doc.add_paragraph("Subcontractors Present:", style="Heading 2")
    for s in subcontractors:
        doc.add_paragraph(f"    - {s}")

    doc.add_paragraph("Areas of Work/Progress:", style="Heading 2")
    for a in areas:
        doc.add_paragraph(f"    - {a}")

    doc.add_page_break()

    for img in images:
        doc.add_picture(img, width=Inches(5.93), height=Inches(4.45))
        doc.paragraphs[-1].alignment = 1  # Center

    # Save to BytesIO
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# === Streamlit App ===
st.title("📷 Site Observation Report Generator")

weather = st.text_input("Weather Conditions (°C)")
max_dim = st.number_input("Max image dimension (px)", min_value=500, max_value=5000, value=1300)
quality = st.slider("JPEG Quality (1 = small file, 100 = high quality)", 1, 100, 95)

subcontractors = st.text_area("Subcontractors (one per line)").split("\n")
areas = st.text_area("Areas of Work (one per line)").split("\n")

uploaded_files = st.file_uploader("Upload Site Images", type=['jpg', 'jpeg'], accept_multiple_files=True)

if st.button("Generate Report"):
    if not uploaded_files:
        st.warning("Please upload at least one image.")
    else:
        with st.spinner("Generating report..."):
            # Sort images by number in filename
            uploaded_files.sort(key=lambda f: extract_number(f.name))

            temp_images = []
            for file in uploaded_files:
                # Use original image (no compression)
                img = resize_image(file.read(), max_dim)
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
                img.save(temp_file.name, format="JPEG", quality=quality, optimize=True)
                temp_images.append(temp_file.name)

            report_bytes = generate_report(temp_images, weather, subcontractors, areas, max_dim, quality)
            st.success("Report generated!")
            st.download_button("📄 Download Report", report_bytes, file_name="Progress_Report.docx")
