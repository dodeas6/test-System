import streamlit as st
import pandas as pd
import re
import io
from openpyxl import load_workbook

# 1. الإعدادات الرسمية
LOGO_URL = "https://i.pinimg.com/originals/5a/65/ee/5a65ee278cd557143f05a4ba91abbfa8.gif"
DEEP_GREEN_BG = "#021a0d"  
IA_GREEN = "#1d4c2b"
BUTTON_GREEN = "#16a34a" 

st.set_page_config(page_title="Iraqi Airways System", layout="wide")

# 2. هندسة التصميم (CSS)
st.markdown(f"""
    <style>
    @import url('https://googleapis.com');

    .stApp {{ background-color: {DEEP_GREEN_BG} !important; }}
    .stMarkdown, p, label, h3, h2, h1, span {{ color: white !important; font-family: 'Cairo', sans-serif; }}
    
    .header-content div, .header-content span {{ color: #1d4c2b !important; font-weight: bold !important; display: block !important; }}

    [data-testid="stFileUploader"] {{ background-color: {BUTTON_GREEN} !important; border-radius: 12px !important; padding: 15px !important; }}
    [data-testid="stFileUploader"] * {{ color: black !important; font-weight: 900 !important; }}

    .stTextArea textarea {{ background-color: #dcfce7 !important; border: 2px solid #4ade80 !important; border-radius: 15px !important; color: #013220 !important; font-weight: bold; }}
    
    /* توحيد تنسيق جميع الأزرار بما فيها زر التحميل */
    .stButton>button, .stDownloadButton>button {{ 
        width: 100% !important; 
        background-color: {BUTTON_GREEN} !important; 
        color: white !important; 
        border-radius: 12px !important; 
        font-weight: 900 !important; 
        padding: 15px !important; 
        border: none !important;
        font-family: 'Cairo', sans-serif !important;
        font-size: 18px !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# 3. نظام تسجيل الدخول
if "auth" not in st.session_state:
    st.session_state["auth"] = False

if not st.session_state["auth"]:
    st.markdown(f'<div style="text-align:center; padding-top:50px;"><img src="{LOGO_URL}" width="250"></div>', unsafe_allow_html=True)
    _, col_mid, _ = st.columns([1, 2, 1])
    with col_mid:
        st.markdown(f"<h2 style='text-align:center;'>بوابة دخول الموظفين</h2>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='text-align:center;'>تصميم المهندس محمد عبد الستار</h2>", unsafe_allow_html=True)
		user = st.text_input("اسم المستخدم")
        pw = st.text_input("كلمة المرور", type="password")
        if st.button("دخول"):
            if user == "iraqi_admin" and pw == "ia2024":
                st.session_state["auth"] = True
                st.rerun()
            else:
                st.error("خطأ في البيانات")
    st.stop()

# 4. الهيدر
st.markdown(f"""
    <div style="background-color:white; padding:30px; border-radius:0 0 35px 35px; text-align:center; box-shadow: 0 10px 20px rgba(0,0,0,0.5);" class="header-content">
        <center><img src="{LOGO_URL}" width="180"></center>
        <div style="font-size:32px; margin-top:15px; font-family:Cairo; color:#1d4c2b !important; font-weight:900;">نظام إدارة بيانات المسافرين والحقائب</div>
        <div style="font-family:Orbitron; font-size:24px; letter-spacing:4px; margin-top:5px; color:#1d4c2b !important; font-weight:bold;">IRAQI AIRWAYS</div>
	<div style="font-family:Orbitron; font-size:24px; letter-spacing:4px; margin-top:5px; color:#1d4c2b !important; font-weight:bold;">Desgined by Eng. Mohammad Abdulsattar </div>
    </div>
    <br>
    """, unsafe_allow_html=True)

template_file = st.file_uploader("📄 الخطوة الأولى: Upload template.xlsx", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    p_input = st.text_area("📋 الصق بيانات المانيفست هنا:", height=280)
with c2:
    b_input = st.text_area("👜 الصق بيانات الحقائب هنا:", height=280)

def extract_val(text, start_key):
    pattern = rf"{start_key}-(.*?)(?=/|$)"
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def find_seat(text):
    match = re.search(r'\b0*(\d{1,3}[A-F])\b', text, re.IGNORECASE)
    return match.group(1).upper() if match else None

def parse_bags(text):
    pcs = re.search(r'(\d+)\s*PCS', text, re.IGNORECASE)
    kg = re.search(r'(\d+)\s*KG', text, re.IGNORECASE)
    return (pcs.group(1) if pcs else "0"), (kg.group(1) if kg else "0")

if st.button("🚀 معالجة البيانات وتوليد التقرير"):
    if not template_file:
        st.error("يرجى رفع ملف template.xlsx")
    elif p_input and b_input:
        p_sections = re.split(r'(\d+\.)', p_input)
        passengers = {}
        for i in range(1, len(p_sections), 2):
            entry = p_sections[i] + p_sections[i+1]
            seat = find_seat(entry)
            if seat:
                passengers[seat] = {
                    "first": extract_val(entry, "FIRST NAME"),
                    "last": extract_val(entry, "SURNAME"),
                    "nat": extract_val(entry, "NATIONALITY"),
                    "gender": extract_val(entry, "GENDER"),
                    "passport": extract_val(entry, "NUMBER"),
                    "seat": seat, "pcs": "0", "kg": "0"
                }

        for row in b_input.split('\n'):
            b_seat = find_seat(row)
            if b_seat and b_seat in passengers:
                pcs, kg = parse_bags(row)
                passengers[b_seat]["pcs"], passengers[b_seat]["kg"] = pcs, kg

        if passengers:
            st.success(f"✅ تمت مطابقة {len(passengers)} مسافر.")
            st.dataframe(pd.DataFrame(list(passengers.values())), use_container_width=True)
            
            wb = load_workbook(template_file)
            ws = wb.active
            passengers_list = list(passengers.values())
            
            for i, data in enumerate(passengers_list):
                curr_row = 2 + i 
                ws.cell(row=curr_row, column=1).value = data["first"]
                ws.cell(row=curr_row, column=2).value = data["last"]
                ws.cell(row=curr_row, column=3).value = data["nat"]
                ws.cell(row=curr_row, column=4).value = data["gender"]
                ws.cell(row=curr_row, column=5).value = "PASSPORT"
                ws.cell(row=curr_row, column=6).value = data["passport"]
                ws.cell(row=curr_row, column=7).value = data["seat"]
                ws.cell(row=curr_row, column=8).value = data["pcs"]
                ws.cell(row=curr_row, column=9).value = data["kg"]

            out = io.BytesIO()
            wb.save(out)
            st.download_button("📥 تحميل المانيفست المكتمل", out.getvalue(), "Iraqi_Airways_Final_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
