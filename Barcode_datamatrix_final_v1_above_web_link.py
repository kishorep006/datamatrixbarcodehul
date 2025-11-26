# -*- coding: utf-8 -*-
"""
Created on Wed Nov 26 16:12:35 2025
@author: kisho
"""

import streamlit as st
import pandas as pd
import os
import zipfile
from pystrich.datamatrix import DataMatrixEncoder
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import tempfile

# -------------------------
# App Title
# -------------------------
st.title("DataMatrix Barcode Generator")
st.write("Upload an Excel file with a column named *Barcode_id* to generate barcodes.")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:

    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Create zip memory buffer
    output_zip = BytesIO()
    with zipfile.ZipFile(output_zip, "w") as zipf:

        mm_to_inch = 1 / 25.4
        target_size_px = int(50 * mm_to_inch * 300)  # 50mm @ 300 DPI

        # -----------------------------
        # Load BOLD TTF font (must be in repo)
        # -----------------------------
        FONT_PATH = "DejaVuSans-Bold.ttf"

        try:
            font = ImageFont.truetype(FONT_PATH, 40)  # Bigger, bold, high DPI
        except Exception as e:
            st.error(f"Font error: {e}")
            font = ImageFont.load_default()

        for i, row in df.iterrows():

            code = str(row["Barcode_id"]).strip()
            if not code:
                continue

            # ---- Temporary file for DMX ----
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_file:
                temp_filename = tmp_file.name

            # Generate DataMatrix
            encoder = DataMatrixEncoder(code)
            encoder.save(temp_filename)

            # Load the DMX image
            barcode_img = Image.open(temp_filename).convert("RGB")
            barcode_img = barcode_img.resize((target_size_px, target_size_px - 120), Image.LANCZOS)

            # Remove temp file
            os.remove(temp_filename)

            # Final white canvas
            final_img = Image.new("RGB", (target_size_px, target_size_px), "white")
            draw = ImageDraw.Draw(final_img)

            # ---- Draw the bold text ----
            # Compute text size (center align)
            try:
                bbox = draw.textbbox((0, 0), code, font=font)
                text_w = bbox[2] - bbox[0]
                text_h = bbox[3] - bbox[1]
            except AttributeError:
                text_w, text_h = draw.textsize(code, font=font)

            text_x = (target_size_px - text_w) // 2
            text_y = 20  # gap from top

            draw.text((text_x, text_y), code, fill="black", font=font)

            # ---- Paste DataMatrix under text ----
            barcode_y = text_y + text_h + 20
            final_img.paste(barcode_img, (0, barcode_y))

            # Save PNG into ZIP buffer
            img_bytes = BytesIO()
            final_img.save(img_bytes, format="PNG", dpi=(300, 300))
            img_bytes.seek(0)

            zipf.writestr(f"{code}.png", img_bytes.read())

    output_zip.seek(0)
    st.success("Barcodes generated successfully!")
    st.download_button("Download ZIP", data=output_zip, file_name="barcodes.zip")
