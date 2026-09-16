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

# ฟังก์ชันป้องกันไม่ให้เผลอดึงหัวข้อตารางมาเป็นข้อมูล
def is_header(text):
    clean_text = text.replace(" ", "")
    headers = [
        "หน่วยงาน:", "หน่วยงาน", "ที่อยู่:", "ที่อยู่", "ประเภทสถานีวิทยุคมนาคม", 
        "เลขที่ใบอนุญาตตั้ง", "เลขที่ใบอนุญาตใช้", "ที่ตั้ง", "ตำบล", "อำเภอ", 
        "จังหวัด", "รหัสไปรษณีย์", "Longtitude", "Longitude", "Latitude", "หมู่ที่:", "หมู่ที่"
    ]
    return clean_text in headers

def find_idx(row, keywords):
    for i, cell in enumerate(row):
        cell_clean = cell.replace(" ", "")
        for k in keywords:
            if k.replace(" ", "") in cell_clean:
                return i
    return -1

# ฟังก์ชันค้นหาข้อมูลแบบ Smart Lookup (มองขวา ถ้าไม่มีให้มองล่าง)
def get_val_in_block(block, keywords):
    for r_idx, row in enumerate(block):
        idx = find_idx(row, keywords)
        if idx != -1:
            # 1. หาในแถวเดียวกัน (คอลัมน์ถัดไปทางขวา)
            for j in range(idx + 1, len(row)):
                val = row[j].strip()
                if val and not is_header(val):
                    return val
            
            # 2. ถ้าแถวเดียวกันไม่มี ให้หาในแถวถัดไป (บรรทัดล่าง)
            if r_idx + 1 < len(block):
                next_row = block[r_idx + 1]
                # ลองดึงจากคอลัมน์ที่ตรงกันในบรรทัดล่าง
                if idx < len(next_row):
                    val = next_row[idx].strip()
                    if val and not is_header(val):
                        return val
                # ถ้าไม่เจอ ลองหาคำแรกในบรรทัดล่างที่ไม่ใช่ Header
                for cell in next_row:
                    val = cell.strip()
                    if val and not is_header(val):
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
                
                # ตัดขึ้นสถานีใหม่เมื่อเจอคำว่า "หน่วยงาน" อย่างเดียว
                blocks = []
                current_block = []
                for row in clean_rows:
                    if find_idx(row, ['หน่วยงาน']) != -1:
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
                    # ดึงข้อมูลด้วยระบบ Smart Lookup
                    data = {
                        'operator': get_val_in_block(block, ['หน่วยงาน']),
                        'license': get_val_in_block(block, ['เลขที่ใบอนุญาตตั้ง']),
                        'location': get_val_in_block(block, ['ที่ตั้ง']),
                        'subdistrict': get_val_in_block(block, ['ตำบล']),
                        'district': get_val_in_block(block, ['อำเภอ']),
                        'province': get_val_in_block(block, ['จังหวัด']),
                        'zipcode': get_val_in_block(block, ['รหัสไปรษณีย์']),
                        'lon': get_val_in_block(block, ['Longitude', 'Longtitude']),
                        'lat': get_val_in_block(block, ['Latitude']),
                        'date_calc': get_val_in_block(block, ['วันที่วัด/คำนวณ']),
                        'signature': get_val_in_block(block, ['ผู้มีอำนาจลงนาม']),
                        'date_report': get_val_in_block(block, ['วันที่รายงาน']),
                        'max_dist_text': "", 'max_dist_val': "", 'max_rad': ""
                    }
                    
                    freq_rows = []
                    
                    for row in block:
                        # สกัดค่าระดับสูงสุด 
                        if find_idx(row, ['ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด']) != -1 and '=' not in "".join(row):
                            data['max_dist_text'] = 'ระดับสูงสุด'
                            vals = [x for x in row if x.strip()]
                            if len(vals) >= 3:
                                data['max_dist_val'] = vals[1]
                                data['max_rad'] = vals[2]
                            elif len(vals) == 2:
                                data['max_dist_val'] = vals[1]
                                
                        # สกัดข้อมูลตารางความถี่
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
