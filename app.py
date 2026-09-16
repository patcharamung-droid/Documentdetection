import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="กสทช. PDF to Excel", page_icon="📶", layout="wide")
st.title("📶 ระบบสกัดข้อมูลรายงาน กสทช. เป็น Excel")
st.write("อัปโหลดไฟล์ PDF แบบรายงานระดับการแผ่คลื่นแม่เหล็กไฟฟ้า ระบบจะจัดเรียงข้อมูลใหม่เป็น 22 คอลัมน์ (1 สถานี ต่อ 1 แถว)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF ของคุณ", type=["pdf"])

# ฟังก์ชันช่วยหาตำแหน่งของคำในแถว
def find_idx(row, keywords):
    for i, cell in enumerate(row):
        if any(k in cell for k in keywords):
            return i
    return -1

# ฟังก์ชันดึงค่าที่อยู่ถัดจากคำค้นหา
def get_next_val(row, keywords):
    idx = find_idx(row, keywords)
    if idx != -1:
        for j in range(idx + 1, len(row)):
            val = row[j].strip()
            if val and not any(k in val for k in keywords):
                return val
    return ""

# ฟังก์ชันยุบรวมข้อมูลความถี่ (ถ้าเหมือนกันหมดจะแสดงค่าเดียว ถ้าต่างกันจะคั่นด้วยลูกน้ำ)
def clean_join(seq):
    if not seq: return ""
    seq = [str(x).strip() for x in seq if str(x).strip()] # ล้างค่าว่าง
    if not seq: return ""
    if len(set(seq)) == 1: # ถ้าข้อมูลทุกบรรทัดเหมือนกัน (เช่น ความสูง 45.0)
        return seq[0]
    return ", ".join(seq) # ถ้าต่างกัน ให้คั่นด้วยลูกน้ำ

if uploaded_file is not None:
    if st.button("เริ่มสกัดข้อมูล"):
        with st.spinner("กำลังประมวลผลเอกสาร..."):
            all_rows = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            if not table: continue
                            all_rows.extend(table)
            
            if all_rows:
                # 1. คลีนข้อมูลและตัดค่า None
                clean_rows = []
                for row in all_rows:
                    cleaned = [str(x).replace('\n', ' ').strip() if x is not None else "" for x in row]
                    if any(cleaned):
                        clean_rows.append(cleaned)
                
                # 2. แยกบล็อกข้อมูลทีละสถานี
                blocks = []
                current_block = []
                for row in clean_rows:
                    if find_idx(row, ['หน่วยงาน:']) != -1:
                        if current_block:
                            blocks.append(current_block)
                        current_block = [row]
                    else:
                        current_block.append(row)
                if current_block:
                    blocks.append(current_block)

                # 3. เริ่มสกัดข้อมูลรายสถานี
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
                        if find_idx(row, ['หน่วยงาน:']) != -1:
                            data['operator'] = get_next_val(row, ['หน่วยงาน:'])
                        elif find_idx(row, ['เลขที่ใบอนุญาตต้งั', 'เลขที่ใบอนุญาตตั้ง']) != -1:
                            data['license'] = get_next_val(row, ['เลขที่ใบอนุญาตต้งั', 'เลขที่ใบอนุญาตตั้ง'])
                        elif find_idx(row, ['ที่ต้งั', 'ที่ตั้ง']) != -1:
                            data['location'] = get_next_val(row, ['ที่ต้งั', 'ที่ตั้ง'])
                        elif find_idx(row, ['ต าบล', 'ตำบล']) != -1:
                            data['subdistrict'] = get_next_val(row, ['ต าบล', 'ตำบล'])
                            data['district'] = get_next_val(row, ['อ าเภอ', 'อำเภอ'])
                        elif find_idx(row, ['จังหวัด']) != -1:
                            data['province'] = get_next_val(row, ['จังหวัด'])
                            data['zipcode'] = get_next_val(row, ['รหัสไปรษณยี ์', 'รหัสไปรษณีย์'])
                        elif find_idx(row, ['Longitude']) != -1:
                            l_idx = find_idx(row, ['Longitude'])
                            if l_idx != -1:
                                for j in range(l_idx + 1, len(row)):
                                    if row[j].strip() and row[j] != 'Longitude':
                                        data['lon'] = row[j]
                                        break
                            lat_idx = find_idx(row, ['Latitude'])
                            if lat_idx != -1:
                                for j in range(lat_idx + 1, len(row)):
                                    if row[j].strip():
                                        data['lat'] = row[j]
                                        break
                                        
                        # สกัดค่าระดับสูงสุด 
                        elif find_idx(row, ['ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด']) != -1 and '=' not in "".join(row):
                            data['max_dist_text'] = 'ระดับสูงสุด'
                            vals = [x for x in row if x.strip()]
                            if len(vals) >= 3:
                                data['max_dist_val'] = vals[1]
                                data['max_rad'] = vals[2]
                                
                        elif find_idx(row, ['วนั ที่วดั /คา นวณ', 'วันที่วัด/คำนวณ']) != -1:
                            data['date_calc'] = get_next_val(row, ['วนั ที่วดั /คา นวณ', 'วันที่วัด/คำนวณ'])
                        elif find_idx(row, ['ผมู้ ีอา นาจลงนาม', 'ผู้มีอำนาจลงนาม']) != -1:
                            data['signature'] = get_next_val(row, ['ผมู้ ีอา นาจลงนาม', 'ผู้มีอำนาจลงนาม'])
                        elif find_idx(row, ['วนั ที่รายงาน', 'วันที่รายงาน']) != -1:
                            data['date_report'] = get_next_val(row, ['วนั ที่รายงาน', 'วันที่รายงาน'])
                            
                        # เก็บข้อมูลตารางความถี่
                        non_empty = [x for x in row if x.strip()]
                        if non_empty and re.match(r'^\d+$', non_empty[0]) and "เมตร" not in non_empty[0] and len(non_empty) >= 5:
                            freq_rows.append(non_empty)

                    # 4. ประกอบร่างข้อมูล 1 สถานี ต่อ 1 แถว (ยุบข้อมูลความถี่)
                    if freq_rows:
                        freqs = clean_join([f[0] for f in freq_rows if len(f) > 0])
                        brands = clean_join([f[1] for f in freq_rows if len(f) > 1])
                        models = clean_join([f[2] for f in freq_rows if len(f) > 2])
                        powers = clean_join([f[3] for f in freq_rows if len(f) > 3])
                        gains = clean_join([f[4] for f in freq_rows if len(f) > 4])
                        heights = clean_join([f[5] for f in freq_rows if len(f) > 5])
                    else:
                        freqs = brands = models = powers = gains = heights = ""

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
                        'ความถี่': freqs,
                        'ตราอักษร': brands,
                        'รุ่น/แบบ': models,
                        'กำลังส่ง (วัตต์)': powers,
                        'อัตราขยายสายอากาศ (dBi)': gains,
                        'ความสูงสายอากาศ (เมตร)': heights,
                        'ระยะห่างจากเสา ที่ต้ังสายอากาศ': data['max_dist_text'],
                        'ระยะที่วัด/คำนวณ (เมตร)': data['max_dist_val'],
                        'ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด': data['max_rad'],
                        'วันที่วัด/คำนวณ': data['date_calc'],
                        'ลงชื่อ': data['signature'],
                        'วันที่รายงาน': data['date_report']
                    }
                    parsed_data.append(record)
                    row_index += 1
                
                # 5. สรุปเป็นไฟล์ Excel
                if parsed_data:
                    final_df = pd.DataFrame(parsed_data)
                    st.dataframe(final_df.head(10)) 
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='Report_Data')
                    processed_data = output.getvalue()
                    
                    st.success(f"สกัดข้อมูลสำเร็จ! สรุปข้อมูลทั้งหมด {len(final_df)} สถานี 🎉")
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ Excel",
                        data=processed_data,
                        file_name="NBTC_Report_Summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("ไม่พบข้อมูลสถานีในแบบฟอร์มครับ")
            else:
                st.warning("ไม่พบโครงสร้างตารางในไฟล์ PDF ครับ")
