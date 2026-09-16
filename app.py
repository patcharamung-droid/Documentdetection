import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄")

st.title("📄 ระบบแปลงตาราง PDF เป็น Excel")
st.write("อัปโหลดไฟล์ PDF ที่มีตาราง ระบบจะดึงข้อมูลทั้งหมดมารวมเป็นไฟล์ .xlsx ให้ทันที")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF ของคุณ", type=["pdf"])

if uploaded_file is not None:
    if st.button("เริ่มดึงข้อมูล"):
        with st.spinner("กำลังอ่านข้อมูลจาก PDF..."):
            all_rows = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            if not table: continue
                            # เปลี่ยนมาใช้วิธีนำ "ข้อมูลทุกแถว" ของตารางนี้ไปเก็บรวมใน List 
                            # เพื่อหลีกเลี่ยง Error จาก pd.concat()
                            all_rows.extend(table)
            
            if all_rows:
                # 1. สร้าง DataFrame จากข้อมูลทุกแถวรวดเดียว (ยังไม่มีชื่อคอลัมน์)
                df = pd.DataFrame(all_rows)
                
                # 2. นำข้อมูลแถวแรก (Index 0) มาเตรียมเป็นหัวตาราง
                # หากช่องไหนว่างเปล่า (None) จะถูกแทนที่ด้วยคำว่า Unnamed_ตามด้วยตัวเลข
                raw_columns = df.iloc[0].tolist()
                clean_columns = [str(c) if c else f"Unnamed_{i}" for i, c in enumerate(raw_columns)]
                
                # 3. จัดการกรณีที่ชื่อคอลัมน์ซ้ำกัน (เช่น มีคอลัมน์ "หมายเหตุ" 2 ช่อง)
                unique_columns = []
                seen = set()
                for col in clean_columns:
                    new_col = col
                    count = 1
                    while new_col in seen:
                        new_col = f"{col}_{count}"
                        count += 1
                    seen.add(new_col)
                    unique_columns.append(new_col)
                    
                # นำชื่อคอลัมน์ที่ไม่ซ้ำกันไปตั้งเป็น Header
                df.columns = unique_columns
                
                # 4. ตัดข้อมูลแถวแรกทิ้ง (เพราะถูกนำไปเป็น Header แล้ว)
                final_df = df[1:].reset_index(drop=True)
                
                # (เสริมความเนียน) ลบแถวที่เป็น "หัวตารางซ้ำ" ที่อาจติดมาจากหน้าที่ 2 เป็นต้นไป
                first_col_name = final_df.columns[0]
                final_df = final_df[final_df[first_col_name] != raw_columns[0]]
                
                # 5. คลีนข้อมูลตัวอักษรขึ้นบรรทัดใหม่
                final_df = final_df.replace('\n', ' ', regex=True)
                
                # เตรียมไฟล์สำหรับดาวน์โหลด
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Extracted_Data')
                processed_data = output.getvalue()
                
                st.success(f"ดึงข้อมูลเสร็จสมบูรณ์! พบข้อมูลทั้งหมด {len(final_df)} แถว 🎉")
                
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel",
                    data=processed_data,
                    file_name="extracted_tables.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("ไม่พบโครงสร้างตารางในไฟล์ PDF ที่อัปโหลดครับ")
