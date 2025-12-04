import streamlit as st
import pandas as pd
import io

st.title("Category-wise Seat Distribution – Single Sheet Generator")

# ===== CATEGORY MAP =====
CATEGORY_MAP = {
    "SM": "SM",
    "EW": "EWS",
    "SC": "SC",
    "ST": "ST",
    "EZ": "EZ",
    "MU": "MU",
    "BH": "BH",
    "LA": "LA",
    "BX": "BX",
    "KU": "KU",
    "DV": "SQ",
    "KN": "SQ"
}

MAIN_COLUMNS = [
    "College", "Course", "Seats",
    "SQ", "SM", "EWS", "SC", "ST", "EZ", "MU",
    "BH", "LA", "BX", "KU",
    "SM-PD", "EWS-PD", "SC-PD", "ST-PD", "EZ-PD",
    "MU-PD", "BH-PD", "LA-PD", "BX-PD"
]

st.write("Upload the raw seat matrix Excel file:")

uploaded_file = st.file_uploader("Choose Excel File", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    # Remove footer rows
    df = df[df["Program"].notna()]

    rows = []
    for _, r in df.iterrows():
        row = {"College": r["College"], "Course": r["Specialty"]}

        # Map categories
        for old, new in CATEGORY_MAP.items():
            row[new] = row.get(new, 0) + r.get(old, 0)

        # PD fields = 0
        for col in ["SM-PD","EWS-PD","SC-PD","ST-PD","EZ-PD",
                    "MU-PD","BH-PD","LA-PD","BX-PD"]:
            row[col] = 0

        # Total seats
        row["Seats"] = sum(row.get(c, 0) for c in
            ["SQ","SM","EWS","SC","ST","EZ","MU","BH","LA","BX","KU"]
        )

        rows.append(row)

    result = pd.DataFrame(rows)

    # Ensure all columns exist
    for col in MAIN_COLUMNS:
        if col not in result.columns:
            result[col] = 0

    # Reorder
    result = result[MAIN_COLUMNS]

    st.success("Generated Successfully!")

    st.dataframe(result)

    # ===== Excel Export =====
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="Seat_Distribution", index=False)

    st.download_button(
        label="📥 Download Final Excel (Single Sheet)",
        data=buffer,
        file_name="Seat_Distribution.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
