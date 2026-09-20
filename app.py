import streamlit as st
import pandas as pd
from datetime import datetime, date

# --- डेटा क्लिनिंग आणि फॉरमॅट फंक्शन्स ---
def get_clean_number(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    if val_str.lower() in ['nan', 'nat', 'none', '']:
        return ""
    return val_str

def format_to_ddmmyyyy(date_val):
    if pd.isna(date_val) or not date_val or str(date_val).lower() in ['nan', 'nat', 'none', '']:
        return ""
    val_str = str(date_val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    if ' ' in val_str:
        val_str = val_str.split(' ')[0]
    try:
        dt = pd.to_datetime(val_str, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.strftime('%d/%m/%Y')
        return val_str
    except Exception:
        return val_str

# --- डमी डेटाबेस कनेक्शन किंवा सि्युलेशन (तुमच्या प्रोजेक्टनुसार युज करा) ---
# येथे समजा `conn` आणि `c` आधीपासून कनेक्ट आहेत.
# import sqlite3
# conn = sqlite3.connect('tapal.db')
# c = conn.cursor()

st.title("📁 टपाल नोंदणी प्रणाली (Tapal Management System)")

# टॅब्स तयार करणे
tab1, tab2 = st.tabs(["➕ नवीन टपाल नोंद", "📋 टपाल यादी / रजिस्टर"])

with tab1:
    st.subheader("नवीन टपाल माहिती भरा")
    
    with st.form("tapal_entry_form"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            file_no = st.text_input("फाइल क्र. (File No)")
            inward_no = st.text_input("आवक क्रमांक (Inward No)")
            inward_date_val = st.date_input("आवक दिनांक (Inward Date)", value=date.today())
            
        with col2:
            computer_no = st.text_input("संगणक क्रमांक (Computer No)")
            subject_code = st.text_input("विषय संकेत (Subject Code)")
            letter_no_date = st.text_input("पत्राचा क्रमांक व दिनांक")
            
        with col3:
            letter_from = st.text_input("पत्र पाठवणारे (Letter From)")
            letter_type = st.text_input("पत्राचा प्रकार")
            emp_name = st.text_input("कर्मचाऱ्याचे नाव")

        subject = st.text_area("विषय (Subject)")
        
        col4, col5 = st.columns(2)
        with col4:
            address = st.text_input("पत्ता")
            district = st.text_input("जिल्हा")
        with col5:
            computer_no_2 = st.text_input("इतर तपशील / संगणक क्र. २")
            remarks = st.text_input("शेरा (Remarks)")

        submit_btn = st.form_submit_button("✅ टपाल सेव्ह करा")

        if submit_btn:
            # डेटा सुरक्षितपणे क्लिन करणे
            clean_inw = get_clean_number(inward_no)
            clean_file = get_clean_number(file_no)
            clean_comp = get_clean_number(computer_no)
            
            # दिनांक स्ट्रिंग फॉरमॅटमध्ये बदलणे
            inward_date_str = inward_date_val.strftime("%d/%m/%Y") if inward_date_val else ""

            try:
                # टीप: तुमच्या टेबलच्या कॉलम रचनेनुसार येथे डेटा इन्सर्ट होईल
                c.execute('''
                    INSERT INTO tapal_entries (
                        file_no, subject_code, inward_no, computer_no, inward_date,
                        letter_no_date, letter_from, letter_type, subject, emp_name,
                        address, district, action_date, action_taken, computer_no_2,
                        application_close_date, final_action, final_action_details, letter_sent_to,
                        paper_go_in_record, calling_report, report_received_date,
                        info_requested_from, office_letter_no_date, reminder_letter_date, remarks
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    clean_file, subject_code, clean_inw, clean_comp, inward_date_str,
                    letter_no_date, letter_from, letter_type, subject, emp_name,
                    address, district, "", "", computer_no_2,
                    "", "", "", "", "नाही", "नाही", "", "", "", "", remarks
                ))
                conn.commit()
                st.success("🎉 टपाल नोंद यशस्वीरीत्या डेटाबेसमध्ये सेव्ह झाली आहे!")
            except Exception as e:
                st.error(ვეד:=f"डेटा सेव्ह करताना त्रुटी आली: {e}")

with tab2:
    st.subheader("दस्तऐवज रजिस्टर सूची")
    
    # डेटा फेच करून दाखवण्याची पद्धत
    try:
        df = pd.read_sql_query("SELECT * FROM tapal_entries", conn)
        if not df.empty:
            # दिनांकाचे कॉलम फॉरमॅट करणे
            if 'inward_date' in df.columns:
                df['inward_date'] = df['inward_date'].apply(format_to_ddmmyyyy)
            if 'inward_no' in df.columns:
                df['inward_no'] = df['inward_no'].apply(get_clean_number)
                
            st.dataframe(df, use_container_width=True)
        else:
            st.info("सध्या कोणतीही टपाल नोंद उपलब्ध नाही.")
    except Exception as e:
        st.warning(f"डेटा लोड करताना माहिती उपलब्ध नाही किंवा टेबल अजून तयार नाही: {e}")
