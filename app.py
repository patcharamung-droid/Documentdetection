import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄")

st.title("📄 ระบบแปลงตาราง PDF เป็น Excel")
st.write("อัปโหลดไฟล์ PDF ที่มีตาราง ระบบจะดึงข้อมูลทั้งหมดมารวมเป็นไฟล์ .xlsx ให้ทันที")

# 1. สร้างปุ่มอัปโหลดไฟล์
uploaded_file = st.file_uploader("เลือกไฟล์ PDF ของคุณ", type=["pdf"])

if uploaded_file is not None:
    # เมื่อกดปุ่มเริ่มทำงาน
    if st.button("เริ่มดึงข้อมูล"):
        with st.spinner("กำลังอ่านข้อมูลจาก PDF..."):
            all_dfs = []
            
            # 2. อ่าน PDF ที่อัปโหลดเข้ามา
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            if not table: continue
                            # นำข้อมูลมาแปลงเป็น DataFrame โดยให้แถวแรกเป็น Header
                            df = pd.DataFrame(table[1:], columns=table[0])
                            all_dfs.append(df)
            
            # 3. จัดการข้อมูลและสร้างไฟล์ Excel
            if all_dfs:
                final_df = pd.concat(all_dfs, ignore_index=True)
                final_df = final_df.replace('\n', ' ', regex=True) # คลีนข้อมูล
                
                # เขียนไฟล์ Excel ลงในหน่วยความจำ (BytesIO)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Extracted_Data')
                processed_data = output.getvalue()
                
                st.success(f"ดึงข้อมูลเสร็จสมบูรณ์! พบข้อมูลทั้งหมด {len(final_df)} แถว 🎉")
                
                # 4. แสดงปุ่มดาวน์โหลดไฟล์ Excel
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel",
                    data=processed_data,
                    file_name="extracted_tables.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("ไม่พบโครงสร้างตารางในไฟล์ PDF ที่อัปโหลดครับ")
