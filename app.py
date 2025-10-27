import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from io import BytesIO

st.set_page_config(page_title="Casting Error Detector", page_icon="🧮", layout="wide")
st.title("🧮 Casting Error Detector")
st.caption("Detects rounding/casting mismatches in Excel formulas (SUM, +, -), highlights errors, and adds comments with rounded sums.")

uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        wb_formula = openpyxl.load_workbook(uploaded_file, data_only=False)
        wb_values = openpyxl.load_workbook(uploaded_file, data_only=True)
    except Exception as e:
        st.error(f"❌ Error loading workbook: {e}")
        st.stop()

    results = []
    error_cells = {}  # {sheet_name: [(cell_coord, rounded_sum)]}

    for sheet_name in wb_formula.sheetnames:
        sheet_f = wb_formula[sheet_name]
        sheet_v = wb_values[sheet_name]

        for row in sheet_f.iter_rows():
            for cell in row:
                if cell.data_type == "f" and isinstance(cell.value, str):
                    formula = cell.value.strip()

                    # --- SUM() formulas ---
                    if formula.upper().startswith("=SUM("):
                        try:
                            range_part = formula.upper().replace("=SUM(", "").replace(")", "")
                            cell_range = sheet_v[range_part]
                            all_cells = [
                                c.value for row_cells in cell_range for c in row_cells
                                if isinstance(c.value, (int, float))
                            ]
                            if all_cells:
                                actual_sum = round(sum(all_cells), 2)
                                rounded_sum = round(sum(round(x, 2) for x in all_cells), 2)
                                match = round(actual_sum, 2) == round(rounded_sum, 2)

                                if not match:
                                    error_cells.setdefault(sheet_name, []).append((cell.coordinate, rounded_sum))

                                results.append({
                                    "Sheet": sheet_name,
                                    "Cell": cell.coordinate,
                                    "Formula": formula,
                                    "Actual Sum": actual_sum,
                                    "Rounded Sum": rounded_sum,
                                    "Status": "✅ OK" if match else "❌ Casting error detected"
                                })
                        except Exception as e:
                            st.warning(f"⚠️ Error parsing SUM at {cell.coordinate}: {e}")

                    # --- + / - formulas ---
                    elif "+" in formula or "-" in formula:
                        try:
                            expr = formula[1:].replace(" ", "")
                            refs = [part for part in expr.replace("+", "|").replace("-", "|").split("|") if part]

                            ref_values = {}
                            for ref in refs:
                                if "!" in ref:
                                    ref = ref.split("!")[-1]
                                try:
                                    v = sheet_v[ref].value
                                    ref_values[ref] = v
                                except Exception:
                                    ref_values[ref] = 0

                            numeric_values = [v for v in ref_values.values() if isinstance(v, (int, float))]

                            if numeric_values:
                                eval_expr = expr
                                for k, v in ref_values.items():
                                    eval_expr = eval_expr.replace(k, str(v if v is not None else 0))

                                actual_sum = round(eval(eval_expr), 2)
                                rounded_sum = round(sum(round(v, 2) for v in numeric_values), 2)
                                match = round(actual_sum, 2) == round(rounded_sum, 2)

                                if not match:
                                    error_cells.setdefault(sheet_name, []).append((cell.coordinate, rounded_sum))

                                results.append({
                                    "Sheet": sheet_name,
                                    "Cell": cell.coordinate,
                                    "Formula": formula,
                                    "Actual Sum": actual_sum,
                                    "Rounded Sum": rounded_sum,
                                    "Status": "✅ OK" if match else "❌ Casting error detected"
                                })
                        except Exception as e:
                            st.warning(f"⚠️ Error evaluating + / - formula at {cell.coordinate}: {e}")

    # ---- Display results ----
    if results:
        df = pd.DataFrame(results)
        st.subheader("📊 Detected Formulas Summary")

        def highlight_status(val):
            if "❌" in val:
                return "color: red; font-weight: bold;"
            elif "✅" in val:
                return "color: green; font-weight: bold;"
            return ""

        st.dataframe(df.style.map(highlight_status, subset=["Status"]))

        # ---- Highlight & Comment in Excel ----
        yellow_fill = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")

        for sheet_name, cells in error_cells.items():
            sheet = wb_formula[sheet_name]
            for cell_ref, rounded_sum in cells:
                cell = sheet[cell_ref]
                cell.fill = yellow_fill
                comment_text = f"Rounded Sum = {rounded_sum}"
                cell.comment = Comment(comment_text, "Casting Error Detector")

        # ---- Save to BytesIO for download ----
        output = BytesIO()
        wb_formula.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Download Highlighted Excel with Comments",
            data=output,
            file_name="CastingErrorHighlighted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("No formulas found in the entire workbook.")
else:
    st.info("⬆️ Please upload an Excel file to begin.")
