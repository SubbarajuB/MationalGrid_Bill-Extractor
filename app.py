import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="National Grid Multi-PDF Extractor", layout="wide")
st.title("\U0001F4C4 National Grid Bill Extractor for EnergyCAP (Multi-PDF)")

# Function to extract data from one PDF
def extract_data_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()

    # Normalize text
    text = re.sub(r"\s+", " ", text)
    text = text.replace("Adjustment for Changes from Normal Weather", "AdjustmentForChangesFromNormalWeather")

    # Basic fields
    account_match = re.search(r"ACCOUNT NUMBER\s+(\d{5}-\d{5})", text)
    account_number = account_match.group(1) if account_match else "Unknown"
    data = {"Vendor": "National Grid", "Account Number": account_number}

    # Field matchers
    optional_fields = {
        "Service Address": re.search(r"SERVICE FOR\s+(.*?),?\s*ALBANY NY", text),
        "Meter Number": re.search(r"METER NUMBER\s+([A-Z0-9]+)", text),
        "Total Cost": re.search(r"Amount Due \$ *([\d,]+\.\d{2})", text),
        "Paperless Credit": re.search(r"Paperless Billing Credit (-?[\d\.]+)", text),
        "Bill Issue Date": re.search(r"CORRESPONDENCE ADDRESS.*?DATE BILL ISSUED\s+(\w+ \d{1,2}, \d{4})", text),
        "Basic Service Charge": re.search(r"Basic Service Charge.*?(\d{1,3}(?:,\d{3})*(?:\.\d{2}))", text, re.IGNORECASE),
        "Weather Adjustment": re.search(r"AdjustmentForChangesFromNormalWeather [-]?[\$]?(-?\d{1,3}(?:,\d{3})*(?:\.\d{2}))", text, re.IGNORECASE),
        "Delivery Service Adjustment": re.search(r"Delivery Service Adj\(s\).*?therms[^\d-]*(-?\d{1,3}(?:,\d{3})*(?:\.\d{2}))", text, re.IGNORECASE),
        "Tariff Surcharge": re.search(r"Tariff Surcharge[^%]*%\s*(\d{1,3}(?:,\d{3})*(?:\.\d{2}))", text, re.IGNORECASE),
        "Total Delivery Services": re.search(r"Total Delivery Services \$ *(\d{1,3}(?:,\d{3})*(?:\.\d{2}))", text, re.IGNORECASE)
    }

    for key, match in optional_fields.items():
        if match:
            val = match.group(1).replace(",", "")
            if key == "Bill Issue Date":
                try:
                    val = datetime.strptime(val, "%b %d, %Y").strftime("%Y-%m-%d")
                except:
                    val = None
            data[key] = val
        else:
            data[key] = None

    # Service Period
    period_match = re.search(r"(\w+ \d{1,2}, \d{4})\s*(?:to|-)+\s*(\w+ \d{1,2}, \d{4})", text)
    if period_match:
        try:
            start_date = datetime.strptime(period_match.group(1), "%b %d, %Y")
            end_date = datetime.strptime(period_match.group(2), "%b %d, %Y")
            data["Service Period Start"] = start_date.strftime("%Y-%m-%d")
            data["Service Period End"] = end_date.strftime("%Y-%m-%d")
        except:
            data["Service Period Start"] = None
            data["Service Period End"] = None
    else:
        data["Service Period Start"] = None
        data["Service Period End"] = None

    # Dynamic therm charge values
    therm_charge_matches = re.findall(
        r"(Next|Over/Last)\s+(\d+)\s+Therms\s+\d+\.\d+\s+x\s+\d+\s+therms\s+([\d,]+\.\d{2})",
        text,
        re.IGNORECASE
    )
    for label, count, charge in therm_charge_matches:
        column_name = f"{label} {count} Therms"
        data[column_name] = charge.replace(",", "")

    return pd.DataFrame([data])

# Upload multiple PDFs
uploaded_files = st.file_uploader("Upload one or more National Grid bill PDFs", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_dataframes = []
    for file in uploaded_files:
        try:
            df = extract_data_from_pdf(file)
            df["Filename"] = file.name  # Optional: add filename source
            all_dataframes.append(df)
        except Exception as e:
            st.warning(f"Could not process {file.name}: {e}")

    if all_dataframes:
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        st.success(f"\u2705 Extracted data from {len(all_dataframes)} files")
        st.dataframe(combined_df)

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Bills')
            return output.getvalue()

        def to_csv(df):
            return df.to_csv(index=False).encode('utf-8')

        st.download_button(
            label="\U0001F4C5 Download Combined Excel",
            data=to_excel(combined_df),
            file_name="EnergyCAP_Combined.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="\U0001F4C3 Download Combined CSV",
            data=to_csv(combined_df),
            file_name="EnergyCAP_Combined.csv",
            mime="text/csv"
        )
