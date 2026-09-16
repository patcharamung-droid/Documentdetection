import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="กสทช. PDF to Excel", page_icon="📶", layout="wide")

st.title("📶 ระบบสกัดข้อมูลรายงาน กสทช. เป็น Excel")
st.write("อัปโหลดไฟล์ PDF แบบรายงานระดับการแผ่คลื่นแม่เหล็กไฟฟ้า ระบบจะจัดเรียงข้อมูลใหม่เป็น 22 คอลัมน์มาตรฐาน")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF ของคุณ", type=["pdf"])

if uploaded_file is not None:
    if st.button("เริ่มสกัดข้อมูล"):
        with st.spinner("กำลังประมวลผลเอกสาร..."):
            all_rows = []
            
            # 1. ดึงข้อมูลตารางทั้งหมดจาก PDF
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            if not table: continue
                            all_rows.extend(table)
            
            if all_rows:
                # 2. คลีนข้อมูลเบื้องต้น
                clean_rows = []
                for row in all_rows:
                    # แปลงเป็น String ลบช่องว่าง และจัดการค่า None
                    cleaned = [str(x).replace('\n', ' ').strip() if x is not None else "" for x in row]
                    if any(cleaned): # เก็บเฉพาะบรรทัดที่มีข้อมูล
                        clean_rows.append(cleaned)
                
                # 3. แยกข้อมูลออกเป็นชุดๆ (บล็อก) ตามจำนวนสถานี
                blocks = []
                current_block = []
                for row in clean_rows:
                    if 'หน่วยงาน:' in row[0]:
                        if current_block:
                            blocks.append(current_block)
                        current_block = [row]
                    else:
                        current_block.append(row)
                if current_block:
                    blocks.append(current_block)

                # 4. ฟังก์ชันสำหรับสกัดข้อมูล 22 คอลัมน์ จากแต่ละสถานี
                parsed_data = []
                row_index = 1
                
                for block in blocks:
                    data = {
                        'operator': "", 'license': "", 'location': "", 'subdistrict': "", 
                        'district': "", 'province': "", 'zipcode': "", 'lon': "", 'lat': "",
                        'max_dist_text': "", 'max_dist_val': "", 'max_rad': "", 
                        'date_calc': "", 'signature': "", 'date_report': ""
                    }
                    freq_rows = []
                    
                    for row in block:
                        if 'หน่วยงาน:' in row[0]:
                            if len(row) > 1: data['operator'] = row[1]
                        elif 'เลขที่ใบอนุญาตต้งั' in row[0] or 'เลขที่ใบอนุญาตตั้ง' in row[0]:
                            data['license'] = row[-1] if len(row) > 2 else (row[1] if len(row) > 1 else "")
                        elif 'ที่ต้งั' in row[0] or 'ที่ตั้ง' in row[0]:
                            if len(row) > 1: data['location'] = row[1]
                        elif 'ต าบล' in row[0] or 'ตำบล' in row[0]:
                            if len(row) > 1: data['subdistrict'] = row[1]
                            idx = [i for i, x in enumerate(row) if 'อ าเภอ' in x or 'อำเภอ' in x]
                            if idx and len(row) > idx[0] + 1: data['district'] = row[idx[0]+1]
                        elif 'จังหวัด' in row[0]:
                            if len(row) > 1: data['province'] = row[1]
                            idx = [i for i, x in enumerate(row) if 'รหัสไปรษณยี ์' in x or 'รหัสไปรษณีย์' in x]
                            if idx and len(row) > idx[0] + 1: data['zipcode'] = row[idx[0]+1]
                        elif 'Longitude' in row[0]:
                            idx_lon = [i for i, x in enumerate(row) if x == 'Longitude']
                            if idx_lon and len(row) > idx_lon[-1] + 1: data['lon'] = row[idx_lon[-1] + 1]
                            idx_lat = [i for i, x in enumerate(row) if x == 'Latitude']
                            if idx_lat and len(row) > idx_lat[-1] + 1: data['lat'] = row[idx_lat[-1] + 1]
                        
                        # กวาดข้อมูลความถี่ (ถ้าขึ้นต้นด้วยตัวเลขความถี่)
                        elif re.match(r'^\d+$', row[0]) and len(row) >= 5 and "เมตร" not in row[0]:
                            freq_rows.append(row)
                            
                        # กวาดข้อมูลระดับการแผ่คลื่น
                        elif 'ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด' in row[0] and '=' not in row[0]:
                            data['max_dist_text'] = 'ระดับสูงสุด'
                            if len(row) >= 3:
                                data['max_dist_val'] = row[1]
                                data['max_rad'] = row[2]
                            elif len(row) == 2:
                                data['max_dist_val'] = row[1]
                        elif 'วนั ที่วดั /คา นวณ' in row[0] or 'วันที่วัด/คำนวณ' in row[0]:
                             if len(row) > 1: data['date_calc'] = row[1]
                        elif 'ผมู้ ีอา นาจลงนาม' in row[0] or 'ผู้มีอำนาจลงนาม' in row[0]:
                             if len(row) > 1: data['signature'] = row[1]
                        elif 'วนั ที่รายงาน' in row[0] or 'วันที่รายงาน' in row[0]:
                             if len(row) > 1: data['date_report'] = row[1]
                             
                    # จัดเรียงข้อมูลแต่ละความถี่เป็น 1 แถว ใน Excel
                    for freq_row in freq_rows:
                        freq = freq_row[0] if len(freq_row) > 0 else ""
                        brand = freq_row[1] if len(freq_row) > 1 else ""
                        model = freq_row[2] if len(freq_row) > 2 else ""
                        power = freq_row[3] if len(freq_row) > 3 else ""
                        gain = freq_row[4] if len(freq_row) > 4 else ""
                        height = freq_row[5] if len(freq_row) > 5 else ""
                        
                        record = {
                            'ลำดับที่': row_index,
                            'ผู้ประกอบการ': data['operator'],
                            'เลขที่ใบอนุญาตตั้ง': data['license'],
                            'ที่ตั้ง': data['location'],
                            'ตำบล': data['subdistrict'],
                            'อำเภอ': data['district'],
                            'จังหวัด': data['province'],
                            'รหัสไปรษณีย์': data['zipcode'],
                            'Longitude': data['lon'],
                            'Latitude': data['lat'],
                            'ความถี่': freq,
                            'ตราอักษร': brand,
                            'รุ่น/แบบ': model,
                            'กำลังส่ง (วัตต์)': power,
                            'อัตราขยายสายอากาศ (dBi)': gain,
                            'ความสูงสายอากาศ (เมตร)': height,
                            'ระยะห่างจากเสา ที่ต้ังสายอากาศ': data['max_dist_text'],
                            'ระยะที่วัด/คำนวณ (เมตร)': data['max_dist_val'],
                            'ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด': data['max_rad'],
                            'วันที่วัด/คำนวณ': data['date_calc'],
                            'ลงชื่อ': data['signature'],
                            'วันที่รายงาน': data['date_report']
                        }
                        parsed_data.append(record)
                        row_index += 1
                
                # 5. สร้าง DataFrame และไฟล์ Excel
                if parsed_data:
                    final_df = pd.DataFrame(parsed_data)
                    
                    st.dataframe(final_df.head(10)) # โชว์ตัวอย่างข้อมูลบนหน้าเว็บ
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='Report_Data')
                    processed_data = output.getvalue()
                    
                    st.success(f"สกัดข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(final_df)} ชุดความถี่ 🎉")
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ Excel",
                        data=processed_data,
                        file_name="NBTC_Structured_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("ไม่สามารถสกัดข้อมูลตามรูปแบบที่กำหนดได้ครับ")
            else:
                st.warning("ไม่พบโครงสร้างตารางในไฟล์ PDF ครับ")
