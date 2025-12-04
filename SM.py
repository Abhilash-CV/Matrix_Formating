import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import io

st.set_page_config(layout="wide")

st.markdown("<h2 style='text-align:center;'>Category-wise distribution of Seats</h2>", unsafe_allow_html=True)
st.write("")

# ========================================================
# COLLEGE MAPPING (update with your full list)
# ========================================================
college_map = {
    "AMB": "College of Pharmaceutical Sciences, Kannur",
    "KKB": "College of Pharmaceutical Sciences, Kozhikkode",
    "TVB": "College of Pharmaceutical Sciences, Thiruvananthapuram",
    # Add others…
}

uploaded = st.file_uploader("Upload Seat Matrix Excel", type=["xlsx"])

if uploaded:

    df = pd.read_excel(uploaded)
    df = df[df["Program"].notna()]   # remove totals footer

    CATEGORY_MAP = {
        "SM": "SM","EW": "EWS","SC": "SC","ST": "ST",
        "EZ": "EZ","MU": "MU","BH": "BH","LA": "LA",
        "BX": "BX","KU": "KU","DV": "SQ","KN": "SQ"
    }

    MAIN_COLS = [
        "Course","Seats","SQ","SM","EWS","SC","ST",
        "EZ","MU","BH","LA","BX","KU",
        "SM-PD","EWS-PD","SC-PD","ST-PD",
        "EZ-PD","MU-PD","BH-PD","LA-PD","BX-PD"
    ]

    rows = []
    for _, r in df.iterrows():
        row = {"College": r["College"], "Course": r["Specialty"]}

        for old, new in CATEGORY_MAP.items():
            row[new] = r.get(old, 0)

        for pdcol in ["SM-PD","EWS-PD","SC-PD","ST-PD",
                      "EZ-PD","MU-PD","BH-PD","LA-PD","BX-PD"]:
            row[pdcol] = 0

        row["Seats"] = sum(row.get(c,0) for c in
                           ["SQ","SM","EWS","SC","ST",
                            "EZ","MU","BH","LA","BX","KU"])

        rows.append(row)

    result = pd.DataFrame(rows)

    # ========================================================
    # Excel export with FULL formatting
    # ========================================================
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Seat Distribution"

    # Borders
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    # Fonts
    font_normal = Font(name="Times New Roman", size=12)
    font_bold = Font(name="Times New Roman", size=12, bold=True)
    font_white_bold = Font(name="Times New Roman", size=12, bold=True, color="FFFFFF")

    # Fills
    fill_college = PatternFill(fgColor="C65911", fill_type="solid")   # dark orange
    fill_header = PatternFill(fgColor="F4B183", fill_type="solid")    # medium orange
    fill_highlight = PatternFill(fgColor="F8CBAD", fill_type="solid") # light orange

    row_pos = 1

    # ---------------------------- PER COLLEGE BLOCK --------------------------
    for col in result["College"].unique():

        cname = college_map.get(col, col)

        # ---------------- COLLEGE TITLE ROW ----------------
        ws.merge_cells(start_row=row_pos, start_column=1,
                       end_row=row_pos, end_column=len(MAIN_COLS))

        cell = ws.cell(row=row_pos, column=1)
        cell.value = f"{col}: {cname}"
        cell.font = font_white_bold
        cell.fill = fill_college
        cell.alignment = Alignment(horizontal="center")

        row_pos += 1

        # ---------------- HEADER ROW ----------------
        for i, h in enumerate(MAIN_COLS, start=1):
            c = ws.cell(row=row_pos, column=i)
            c.value = h
            c.font = font_bold
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center")
            c.border = border

        row_pos += 1

        block = result[result["College"] == col][MAIN_COLS]

        # add TOTAL row
        total = block.sum(numeric_only=True)
        total["Course"] = "Total"
        block = pd.concat([block, total.to_frame().T], ignore_index=True)

        # ---------------- DATA ROWS ----------------
        for r in dataframe_to_rows(block, index=False, header=False):
            ws.append(r)
            for i, v in enumerate(r, start=1):
                c = ws.cell(row=row_pos, column=i)
                c.border = border
                c.font = font_normal
                c.alignment = Alignment(horizontal="center")

                if isinstance(v,(int,float)) and v > 0:
                    c.fill = fill_highlight
                    c.font = font_bold
            row_pos += 1

        row_pos += 2   # spacing

    # ------------------------- DOWNLOAD BUTTON -------------------------
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)

    st.download_button(
        "📥 Download Excel (FULL Formatted – Screenshot Style)",
        data=excel_buffer.getvalue(),
        file_name="Seat_Distribution_Formatted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

