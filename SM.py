import streamlit as st
import pandas as pd
import numpy as np
import base64
import io

st.set_page_config(layout="wide")

st.markdown("<h2 style='text-align:center;'>Category-wise distribution of the Seats in Government & Private PHARMACY Colleges</h2>", unsafe_allow_html=True)
st.write("")

# ============================
# COLLEGE MAPPING (Paste your mapping here)
# ============================
college_map = {
    "ALB": "College of Pharmaceutical Sciences, Alappuzha",
    "AMB": "College of Pharmaceutical Sciences, Kannur",
    "KKB": "College of Pharmaceutical Sciences, Kozhikkode",
    "KTB": "College of Pharmaceutical Sciences, Kottayam",
    "TVB": "College of Pharmaceutical Sciences, Thiruvananthapuram",
    "AAB": "Al-Azhar College of Pharmacy, Thodupuzha",
    "AFB": "Al Shifa College of Pharmacy, Poonthavanam, Perinthalmanna",
    "AHB": "Ahalia School of Pharmacy, Kozhipara, Palakkad",
    "CAB": "Caritas College of Pharmacy, Ettumanoor, Kottayam",
    "CCB": "Chemists College of Pharmaceutical Sciences & Research, Ernakulam",
    "CHB": "MGM College of Pharmaceutical Sciences, Malappuram",
    "CRB": "Crescent College of Pharmaceutical Sciences, Kannur",
    "DAB": "Devaki Amma Memorial College of Pharmacy, Malappuram",
    "DMB": "Dr. Moopen's College of Pharmacy, Wayanad",
    "DPB": "Department of Pharmaceutical Science, Ettumanoor",
    "DVB": "Dale View College of Pharmacy & Research Centre, TVM",
    "ECB": "Ezhuthachan College of Pharmaceutical Science, TVM",
    "EPB": "Elims College of Pharmacy, Thrissur",
    "GCB": "Grace College of Pharmacy, Palakkad",
    "HGB": "Holy Grace Academy of Pharmacy, Thrissur",
    "HKB": "Hindustan College of Pharmacy, Kottayam",
    "JCB": "St. James College of Pharmaceutical Sciences, Chalakudy",
    "JDB": "JDT Islam College of Pharmacy, Kozhikode",
    "JSB": "Jamia Salafiya Pharmacy College, Malappuram",
    "KAB": "Kerala Academy of Pharmacy, Thiruvananthapuram",
    "KCB": "KMCT College of Pharmacy, Malappuram",
    "KEB": "Indira Gandhi Institute of Pharmaceutical Sciences, Ernakulam",
    "KKP": "KMCT Institute of Pharmaceutical Education & Research, Kozhikode",
    "KLB": "KMCT Institute of Pharmacy, Malappuram",
    "KMB": "KMCT College of Pharmaceutical Sciences, Kozhikode",
    "KNB": "College of Pharmacy, Kannur Medical College",
    "KPB": "KTN College of Pharmacy, Palakkad",
    "KRB": "Karuna College of Pharmacy, Palakkad",
    "KVB": "KVM College of Pharmacy, Cherthala",
    "LIB": "Lisie College of Pharmacy, Ernakulam",
    "MAB": "Malabar College of Pharmacy, Malappuram",
    "MCB": "Moulana College of Pharmacy, Malappuram",
    "MDB": "Mar Dioscorus College of Pharmacy, Thiruvananthapuram",
    "MEB": "MET's College of Pharmaceutical Sciences, Thrissur",
    "MGB": "MGM Silver Jubilee College of Pharmacy, Ernakulam",
    "MGK": "MGM Silver Jubilee College of Pharmacy, Thiruvananthapuram",
    "MGR": "MGM College of Pharmacy, Kannur",
    "MKB": "Malik Deenar College of Pharmacy, Kasaragod",
    "MLB": "Madin College of Pharmacy, Malappuram",
    "MMB": "Mookambika College of Pharmaceutical Sciences, Muvattupuzha",
    "MPB": "Dr Joseph Mar Thoma Institute of Pharmaceutical Sciences, Kattanam",
    "MZB": "Mount Zion College of Pharmaceutical Science, Adoor",
    "NAB": "Nazareth College of Pharmacy, Pathanamthitta",
    "NCB": "Nehru College of Pharmacy, Thrissur",
    "NEB": "Nirmala College of Health Science, Thrissur",
    "NMB": "Nirmala College of Pharmacy, Ernakulam",
    "NPB": "National College of Pharmacy, Kozhikode",
    "PCB": "Pushpagiri College of Pharmacy, Pathanamthitta",
    "PPB": "Prime College of Pharmacy, Palakkad",
    "RGB": "Rajiv Gandhi Institute of Pharmacy, Kasaragod",
    "RIB": "Dept of Pharmaceutical Science, Kottayam",
    "SCB": "St. Joseph's College of Pharmacy, Alappuzha",
    "SJB": "St. John's College of Pharmaceutical Sciences, Idukki",
    "SKP": "Sree Krishna College of Pharmacy, TVM",
    "SPB": "Sanjoe College of Pharmaceutical Studies, Palakkad",
    "STB": "Sree Gokulam SNGM College of Pharmacy, Alappuzha",
    "TIB": "Triveni Institute of Pharmacy, Thrissur",
    "WSB": "Westfort College of Pharmacy, Thrissur"
}

# ============================
# File Upload
# ============================
uploaded = st.file_uploader("Upload Seat Matrix Excel", type=["xlsx"])

if uploaded:
    
    df = pd.read_excel(uploaded)
    df = df[df["Program"].notna()]  # remove footer

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

    MAIN_COLS = [
        "Course", "Seats", "SQ", "SM", "EWS", "SC", "ST", "EZ", "MU",
        "BH", "LA", "BX", "KU",
        "SM-PD", "EWS-PD", "SC-PD", "ST-PD", "EZ-PD",
        "MU-PD", "BH-PD", "LA-PD", "BX-PD"
    ]

    out_rows = []

    for _, r in df.iterrows():
        row = {"Course": r["Specialty"], "College": r["College"]}

        for old, new in CATEGORY_MAP.items():
            row[new] = row.get(new, 0) + r.get(old, 0)

        for col in ["SM-PD","EWS-PD","SC-PD","ST-PD","EZ-PD",
                    "MU-PD","BH-PD","LA-PD","BX-PD"]:
            row[col] = 0

        row["Seats"] = sum(row.get(c, 0) for c in
                           ["SQ","SM","EWS","SC","ST","EZ","MU","BH","LA","BX","KU"])
        out_rows.append(row)

    result = pd.DataFrame(out_rows)
    result = result[["College"] + MAIN_COLS]

    # =============================
    # DISPLAY EXACT FORMAT (SCREEN)
    # =============================

    for college in result["College"].unique():
        block = result[result["College"] == college].copy()

        cname = college_map.get(college, college)

        st.markdown(
            f"<h4 style='text-align:center; margin-top:30px;'>{college} : {cname}</h4>",
            unsafe_allow_html=True
        )

        block = block[MAIN_COLS]
        total = block.sum(numeric_only=True)
        total["Course"] = "Total"

        block = pd.concat([block, total.to_frame().T], ignore_index=True)

        st.dataframe(block, use_container_width=True)

    # GRAND TOTAL
    gtotal = result[MAIN_COLS].sum(numeric_only=True)
    gtotal["Course"] = "Grand Total"

    st.markdown("<h4 style='text-align:center; margin-top:40px;'>Grand Total</h4>", unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([gtotal]), use_container_width=True)
