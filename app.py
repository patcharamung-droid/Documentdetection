import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="กสทช. PDF to Excel", page_icon="📶", layout="wide")
st.title("📶 ระบบสกัดข้อมูลรายงาน กสทช. เป็น Excel")
st.write("อัปโหลดไฟล์ PDF แบบรายงานระดับการแผ่คลื่นแม่เหล็กไฟฟ้า ระบบจะจัดเรียงข้อมูลใหม่เป็น 22 คอลัมน์ (รองรับหลายรูปแบบ)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF ของคุณ", type=["pdf"])

def fix_thai_typos(text):
    if not text: return ""
    typos = {
        "ต าบล": "ตำบล", "อ าเภอ": "อำเภอ", "จา กดั": "จำกัด",
        "ที่ต้งั": "ที่ตั้ง", "หนา้ สา รวจ": "หน้าสำรวจ", "วนั ที่": "วันที่",
        "คา นวณ": "คำนวณ", "วดั ": "วัด", "ผมู้ ีอา นาจ": "ผู้มีอำนาจ",
        "กระทา การ": "กระทำการ", "บริษทั": "บริษัท", "รหัสไปรษณยี ์": "รหัสไปรษณีย์"
    }
    for wrong, right in typos.items():
        text = text.replace(wrong, right)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def find_idx(row, keywords):
    for i, cell in enumerate(row):
        # ลบช่องว่างก่อนเทียบ เพื่อให้รองรับ "หน่วยงาน:" และ "หน่วยงาน :"
        cell_clean = cell.replace(" ", "")
        for k in keywords:
            if k.replace(" ", "") in cell_clean:
                return i
    return -1

def get_next_val(row, keywords):
    idx = find_idx(row, keywords)
    if idx != -1:
        for j in range(idx + 1, len(row)):
            val = row[j].strip()
            # ตรวจสอบเพื่อไม่ให้ดึงค่าที่เป็น Header ซ้ำ
            if val and not any(k.replace(" ", "") in val.replace(" ", "") for k in keywords):
                return val
    return ""

def clean_join(seq):
    if not seq: return ""
    seq = [str(x).strip() for x in seq if str(x).strip()]
    if not seq: return ""
    if len(set(seq)) == 1: 
        return seq[0]
    return ", ".join(seq) 

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
                clean_rows = []
                for row in all_rows:
                    cleaned = [fix_thai_typos(str(x).replace('\n', ' ')) if x is not None else "" for x in row]
                    if any(cleaned):
                        clean_rows.append(cleaned)
                
                blocks = []
                current_block = []
                for row in clean_rows:
                    if find_idx(row, ['หน่วยงาน:', 'หน่วยงาน']) != -1 and 'บริษัท' in "".join(row):
                        if current_block:
                            blocks.append(current_block)
                        current_block = [row]
                    else:
                        current_block.append(row)
                if current_block:
                    blocks.append(current_block)

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
                        if find_idx(row, ['หน่วยงาน']) != -1:
                            data['operator'] = get_next_val(row, ['หน่วยงาน'])
                        elif find_idx(row, ['เลขที่ใบอนุญาตตั้ง']) != -1:
                            data['license'] = get_next_val(row, ['เลขที่ใบอนุญาตตั้ง'])
                        elif find_idx(row, ['ที่ตั้ง']) != -1:
                            data['location'] = get_next_val(row, ['ที่ตั้ง'])
                        elif find_idx(row, ['ตำบล']) != -1:
                            data['subdistrict'] = get_next_val(row, ['ตำบล'])
                            data['district'] = get_next_val(row, ['อำเภอ'])
                        elif find_idx(row, ['จังหวัด']) != -1:
                            data['province'] = get_next_val(row, ['จังหวัด'])
                            data['zipcode'] = get_next_val(row, ['รหัสไปรษณีย์'])
                            
                        # รองรับคำว่า Longitude และ Longtitude
                        elif find_idx(row, ['Longitude', 'Longtitude']) != -1:
                            l_idx = find_idx(row, ['Longitude', 'Longtitude'])
                            if l_idx != -1:
                                for j in range(l_idx + 1, len(row)):
                                    val = row[j].strip()
                                    if val and "Long" not in val and "Lat" not in val:
                                        data['lon'] = val
                                        break
                            lat_idx = find_idx(row, ['Latitude'])
                            if lat_idx != -1:
                                for j in range(lat_idx + 1, len(row)):
                                    val = row[j].strip()
                                    if val and "Lat" not in val:
                                        data['lat'] = val
                                        break
                                        
                        # สกัดค่าระดับสูงสุด 
                        elif find_idx(row, ['ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด']) != -1 and '=' not in "".join(row):
                            data['max_dist_text'] = 'ระดับสูงสุด'
                            vals = [x for x in row if x.strip()]
                            if len(vals) >= 3:
                                data['max_dist_val'] = vals[1]
                                data['max_rad'] = vals[2]
                            elif len(vals) == 2: # เผื่อกรณีคอลัมน์เลื่อน
                                data['max_dist_val'] = vals[1]
                                
                        # รองรับ "วันที่วัด/คำนวณ" และ "วันที่วัด / คำนวณ"
                        elif find_idx(row, ['วันที่วัด/คำนวณ', 'วันที่วัด/คำนวณ']) != -1:
                            data['date_calc'] = get_next_val(row, ['วันที่วัด/คำนวณ'])
                        elif find_idx(row, ['ผู้มีอำนาจลงนาม']) != -1:
                            data['signature'] = get_next_val(row, ['ผู้มีอำนาจลงนาม'])
                        elif find_idx(row, ['วันที่รายงาน']) != -1:
                            data['date_report'] = get_next_val(row, ['วันที่รายงาน'])
                            
                        non_empty = [x for x in row if x.strip()]
                        if non_empty and re.match(r'^\d+$', non_empty[0]) and "เมตร" not in non_empty[0] and len(non_empty) >= 5:
                            freq_rows.append(non_empty)

                    if freq_rows:
                        freqs = clean_join([f[0] for f in freq_rows if len(f) > 0])
                        brands = clean_join([f[1] for f in freq_rows if len(f) > 1])
                        models = clean_join([f[2] for f in freq_rows if len(f) > 2])
                        powers = clean_join([f[3] for f in freq_rows if len(f) > 3])
                        gains = clean_join([f[4] for f in freq_rows if len(f) > 4])
                        heights = clean_join([f[5] for f in freq_rows if len(f) > 5])
                    else:
                        freqs = brands = models = powers = gains = heights = ""

                    if data['operator'] != "" or freqs != "":
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
