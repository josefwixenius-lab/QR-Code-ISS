#!/usr/bin/env python3
"""
QR Code Generator with Customizable Options

Adjust the configuration parameters below as needed.
- QR_FILL_COLOR: Color of the QR code modules (e.g., "black", "#000000").
- QR_BACK_COLOR: Background color (e.g., "white", "#FFFFFF").
- QR_ERROR_CORRECTION: Error correction level (using qrcode.constants.ERROR_CORRECT_*).
- QR_BORDER: Thickness of the border (in boxes).
- QR_VERSION: QR code version; set to None to let the library choose automatically.
"""

import qrcode
import pandas as pd
import urllib.parse

# --- Configuration Parameters ---
QR_FILL_COLOR = "#030037"        # New fill color
QR_BACK_COLOR = "white"          # Background color
QR_ERROR_CORRECTION = qrcode.constants.ERROR_CORRECT_H  # High error correction
QR_BORDER = 4                     # Border size (default is 4 boxes)
QR_VERSION = None                 # QR code version (None = automatic sizing)
# --------------------------------

def generate_qr(url, fill_color, back_color, error_correction, border, version):
    """Generate a QR code image for a given URL."""
    qr = qrcode.QRCode(
        version=version,
        error_correction=error_correction,
        box_size=10,   # Size of each QR code box (pixel size)
        border=border,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color=fill_color, back_color=back_color).convert('RGB')
    return img

def main():
    # Read data from an Excel file (replace with your actual file name/path)
    df = pd.read_excel("data.xlsx")
    df.columns = df.columns.str.strip()  # Trim extra whitespace from column names
    print("Columns after stripping:", df.columns)  # Optional: for debugging

    for index, row in df.iterrows():
        # Extract columns
        amnesrad = str(row["Ämnesrad med ID"]).strip()
        placering = str(row["Placering"]).strip()

        # Build mailto link
        subject_encoded = urllib.parse.quote(amnesrad)
        body_text = f"{amnesrad}\n{placering}\n\nBeskrivning av fel:"
        body_encoded = urllib.parse.quote(body_text)
        mailto_link = (
            f"mailto:servicecenter.eon@se.issworld.com"
            f"?subject={subject_encoded}"
            f"&body={body_encoded}"
        )

        # Generate the QR code
        qr_img = generate_qr(
            url=mailto_link,
            fill_color=QR_FILL_COLOR,
            back_color=QR_BACK_COLOR,
            error_correction=QR_ERROR_CORRECTION,
            border=QR_BORDER,
            version=QR_VERSION
        )

        # Create a safe filename for the QR code image
        safe_filename = amnesrad.lower().replace(" ", "_").replace("/", "_")
        output_filename = f"qr_code_{safe_filename}.png"

        # Save the QR code image
        try:
            qr_img.save(output_filename)
            print(f"QR code for '{amnesrad}' successfully generated and saved as '{output_filename}'.")
        except Exception as e:
            print(f"Error saving QR code image for '{amnesrad}': {e}")

if __name__ == "__main__":
    main()
