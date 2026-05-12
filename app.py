import streamlit as st
import pandas as pd
import re
import io
import os
from openpyxl import load_workbook

# 1. إعدادات الحماية (اليوزر والباسورد)
# يمكنك تغييرها من هنا
USER_ID = "iraqi_admin"
USER_PASS = "ia2024"

# 2. دالة التحقق من الدخول
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        st.markdown("""
            <style>
            .login-box {
                background-color: #ffffff;
                padding: 40px;
                border-radius: 15px;
                box-shadow: 0 4px 20px rgba(0,0,0,0.1);
                border-top: 5px solid #10b981;
            }
            </style>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.markdown('<div class="login-box">', unsafe_allow_html=True)
            st.image("https://wikimedia.org", width=100)
            st.title("تسجيل الدخول")
            user = st.text_input("اسم المستخدم")
            pw = st.text_input("كلمة المرور", type="password")
            if st.button("دخول"):
                if user == USER_ID and pw == USER_PASS:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("❌ البيانات خاطئة")
            st.markdown('</div>', unsafe_allow_html=True)
        return False
    return True

# 3. إعداد الصفحة
st.set_page_config(page_title="Iraqi Airways System", layout="wide")

if check_password():
    # --- هنا يبدأ كود الموقع الأصلي الذي صممناه سابقاً ---
    st.markdown("""
        <style>
        .stApp { background-color: #f0fdf4; }
        .header-box { background: linear-gradient(90deg, #065f46 0%, #10b981 100%); padding: 20px; border-radius: 15px; color: white; text-align: center; margin-bottom:20px; }
        </style>
        <div class="header-box">
            <h1>نظام المانيفست الموحد - الخطوط الجوية العراقية</h1>
            <button style="float:left; border-radius:5px; border:none; padding:5px 10px;" onclick="window.location.reload()">تسجيل خروج</button>
        </div>
    """, unsafe_allow_html=True)

    def extract_val_new(text, start_key):
        pattern = rf"{start_key}-(.*?)(?=/|$)"
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    def find_seat_smart(text):
        match = re.search(r'\b0*(\d{1,3}[A-F])\b', text, re.IGNORECASE)
        return match.group(1).upper() if match else None

    def parse_baggage(text):
        pcs = re.search(r'(\d+)\s*PCS', text, re.IGNORECASE)
        kg = re.search(r'(\d+)\s*KG', text, re.IGNORECASE)
        return (pcs.group(1) if pcs else "0"), (kg.group(1) if kg else "0")

    # --- واجهة العمل ---
    template_file = st.file_uploader("📁 ارفع ملف الفورمة (template.xlsx) أولاً", type=["xlsx"])
    
    col1, col2 = st.columns(2)
    with col1:
        p_input = st.text_area("1️⃣ بيانات المانيفست (SURNAME-)", height=250)
    with col2:
        b_input = st.text_area("2️⃣ بيانات الحقائب (PCS / KG)", height=250)

    if st.button("🚀 مطابقة وتوليد الملف"):
        if not template_file:
            st.error("❌ يرجى رفع ملف الفورمة أولاً!")
        elif p_input and b_input:
            p_sections = re.split(r'(\d+\.)', p_input)
            passengers_dict = {}
            for i in range(1, len(p_sections), 2):
                full_entry = p_sections[i] + p_sections[i+1]
                seat = find_seat_smart(full_entry)
                if seat:
                    passengers_dict[seat] = {
                        "first": extract_val_new(full_entry, "FIRST NAME"),
                        "last": extract_val_new(full_entry, "SURNAME"),
                        "nat": extract_val_new(full_entry, "NATIONALITY"),
                        "gender": extract_val_new(full_entry, "GENDER"),
                        "passport": extract_val_new(full_entry, "NUMBER"),
                        "seat": seat, "pcs": "0", "kg": "0"
                    }

            for b_row in b_input.split('\n'):
                b_seat = find_seat_smart(b_row)
                if b_seat and b_seat in passengers_dict:
                    pcs, kg = parse_baggage(b_row)
                    passengers_dict[b_seat]["pcs"], passengers_dict[b_seat]["kg"] = pcs, kg

            if passengers_dict:
                st.dataframe(pd.DataFrame(passengers_dict.values()))
                wb = load_workbook(template_file)
                ws = wb.active
                for i, data in enumerate(passengers_dict.values()):
                    curr = 2 + i
                    ws.cell(row=curr, column=1).value = data["first"]
                    ws.cell(row=curr, column=2).value = data["last"]
                    ws.cell(row=curr, column=3).value = data["nat"]
                    ws.cell(row=curr, column=4).value = data["gender"]
                    ws.cell(row=curr, column=5).value = "PASSPORT"
                    ws.cell(row=curr, column=6).value = data["passport"]
                    ws.cell(row=curr, column=7).value = data["seat"]
                    ws.cell(row=curr, column=8).value = data["pcs"]
                    ws.cell(row=curr, column=9).value = data["kg"]

                output = io.BytesIO()
                wb.save(output)
                st.success("✅ تمت العملية بنجاح!")
                st.download_button("📥 تحميل المانيفست النهائي", output.getvalue(), "Iraqi_Airways_Final.xlsx")
