import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter

st.set_page_config(page_title="Casting Error Detector", page_icon="🧮", layout="centered")
st.title("🧮 Casting Error Detector")

st.write("Upload an Excel file to check for casting/rounding errors in SUM formulas.")

uploaded_file = st.file_uploader("📂 Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=False)
    sheet = wb.active

    results = []
    counter = 1

    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.upper().startswith("=SUM("):
                # Extract range, e.g. "=SUM(A1:A10)"
                formula = cell.value.strip()
                range_part = formula.replace("=SUM(", "").replace(")", "")
                min_col, min_row, max_col, max_row = range_boundaries(range_part)

                values = []
                for r in range(min_row, max_row + 1):
                    for c in range(min_col, max_col + 1):
                        val = sheet.cell(row=r, column=c).value
                        if isinstance(val, (int, float)):
                            values.append(val)

                # Compute both sums as per your logic
                actual_sum = round(sum(values), 2)  # round final total only
                rounded_sum = round(sum(round(v, 2) for v in values), 2)  # round each first
                matches = abs(actual_sum - rounded_sum) < 1e-9

                results.append({
                    "No.": counter,
                    "Sum Cell": f"{get_column_letter(cell.column)}{cell.row}",
                    "Formula Range": range_part,
                    "Actual Sum": actual_sum,
                    "Rounded Sum": rounded_sum,
                    "Status": "✅" if matches else "❌ Casting error detected"
                })
                counter += 1

    if results:
        df = pd.DataFrame(results)

        # Add color highlighting for status
        def highlight_status(val):
            color = "green" if "✅" in val else "red"
            return f"color: {color}; font-weight: bold;"

        st.dataframe(df.style.applymap(highlight_status, subset=["Status"]))
    else:
        st.warning("No SUM formulas found in the sheet.")
