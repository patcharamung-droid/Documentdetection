import streamlit as st

st.set_page_config(page_title="กสทช. PDF to Excel/CSV", page_icon="📶", layout="wide")

try:
    from extractor import (
        QA_COLUMNS,
        REPORT_COLUMNS,
        build_qa_records,
        build_report_records,
        create_excel_bytes,
        create_csv_bytes,
        parse_pdf_bytes,
    )
except ModuleNotFoundError as error:
    if error.name == "extractor":
        st.error("ไม่พบไฟล์ extractor.py ในโปรเจกต์ที่นำขึ้น Streamlit Cloud")
        st.info("เพิ่มไฟล์ extractor.py ไว้โฟลเดอร์เดียวกับ app.py แล้ว commit และ redeploy อีกครั้ง")
    elif error.name in {"pdfplumber", "pypdf"}:
        st.error(f"ยังไม่ได้ติดตั้งไลบรารี {error.name}")
        st.info("ตรวจว่ามี requirements.txt อยู่โฟลเดอร์เดียวกับ app.py แล้ว commit และ redeploy อีกครั้ง")
    else:
        st.error(f"ไม่พบไลบรารีที่จำเป็น: {error.name}")
    st.stop()
except ImportError:
    st.error("ไฟล์ extractor.py ใน Streamlit Cloud ไม่ตรงกับ app.py เวอร์ชันปัจจุบัน")
    st.info("commit app.py, extractor.py และ requirements.txt เวอร์ชันล่าสุด แล้ว redeploy อีกครั้ง")
    st.stop()

st.title("ระบบสกัดรายงาน กสทช. เป็น Excel หรือ CSV")
st.write(
    "อัปโหลดรายงาน PDF ได้หลายไฟล์ ระบบจะแยกข้อมูลทีละหน้า "
    "จึงรองรับทั้งรายงาน 1 สถานีต่อไฟล์และหลายสถานีในไฟล์เดียว"
)

st.markdown(
    """
    <style>
    div[data-testid="stDownloadButton"] > button {
        width: 100%;
        min-height: 3.2rem;
        border-radius: 12px;
        font-weight: 700;
        font-size: 1rem;
    }
    .download-heading {
        margin-bottom: 0.15rem;
        font-size: 1.1rem;
        font-weight: 700;
    }
    .download-detail {
        color: #5d6570;
        min-height: 3rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.expander("ข้อมูลที่ระบบสกัด", expanded=False):
    st.write(
        "ผู้ประกอบการ ใบอนุญาต ที่ตั้ง พิกัด รายละเอียดความถี่ "
        "ระดับการแผ่คลื่นแม่เหล็กไฟฟ้าสูงสุด วันที่วัด/คำนวณ ผู้ลงนาม และวันที่รายงาน"
    )

uploaded_files = st.file_uploader(
    "เลือกไฟล์รายงาน PDF",
    type=["pdf"],
    accept_multiple_files=True,
    help="แต่ละหน้าควรเป็นแบบรายงานของสถานีเดียว",
)

output_mode = st.radio(
    "รูปแบบข้อมูลที่ต้องการสกัด",
    options=("สรุป 1 แถวต่อสถานี", "ละเอียด 1 แถวต่อความถี่"),
    horizontal=True,
)

pdf_reading_mode = st.radio(
    "วิธีอ่าน PDF",
    options=(
        "อัตโนมัติ (แนะนำ)",
        "อ่านตารางด้วย pdfplumber",
        "อ่านข้อความด้วย pypdf",
    ),
    horizontal=True,
    help=(
        "อัตโนมัติจะอ่านตารางก่อน และใช้ตัวอ่านข้อความสำรองเฉพาะช่องสำคัญที่ยังว่าง "
        "เหมาะกับ PDF ที่ใช้ฟอนต์ไทยพิเศษ"
    ),
)

parser_modes = {
    "อัตโนมัติ (แนะนำ)": "auto",
    "อ่านตารางด้วย pdfplumber": "pdfplumber",
    "อ่านข้อความด้วย pypdf": "pypdf",
}

if st.button("เริ่มสกัดข้อมูล", type="primary", disabled=not uploaded_files):
    all_stations = []
    all_errors = []

    with st.spinner("กำลังอ่านและจัดรูปแบบรายงาน..."):
        for uploaded_file in uploaded_files:
            stations, errors = parse_pdf_bytes(
                uploaded_file.name,
                uploaded_file.getvalue(),
                parser_mode=parser_modes[pdf_reading_mode],
            )
            all_stations.extend(stations)
            all_errors.extend(errors)

    if not all_stations:
        st.error("ไม่พบข้อมูลสถานีที่สกัดได้ โปรดลองตรวจสอบว่า PDF มีตารางข้อความที่เลือกคัดลอกได้")
        if all_errors:
            st.code("\n".join(all_errors), language=None)
    else:
        detail_mode = output_mode == "ละเอียด 1 แถวต่อความถี่"
        report_records = build_report_records(all_stations, detail_mode)
        qa_records = build_qa_records(all_stations)
        report_csv = create_csv_bytes(report_records, REPORT_COLUMNS)
        excel_data = create_excel_bytes(report_records, qa_records)

        needs_review = sum(1 for row in qa_records if row["สถานะ"] == "ต้องตรวจสอบ")
        left, middle, right = st.columns(3)
        left.metric("สถานีที่พบ", len(all_stations))
        middle.metric("แถวใน CSV", len(report_records))
        right.metric("หน้าที่ต้องตรวจสอบ", needs_review)

        st.subheader("ตัวอย่างข้อมูล")
        st.dataframe(report_records[:20], use_container_width=True, hide_index=True)

        if needs_review:
            st.warning(
                "บางหน้าอ่านข้อมูลได้ไม่ครบ โปรดตรวจรายการที่แจ้งว่าไม่พบข้อมูล "
                "และวิธีอ่าน PDF ก่อนนำรายงานไปใช้งาน"
            )
            st.dataframe(
                [row for row in qa_records if row["สถานะ"] == "ต้องตรวจสอบ"],
                use_container_width=True,
                hide_index=True,
            )

        if all_errors:
            with st.expander(f"รายละเอียด {len(all_errors)} หน้าที่อ่านไม่ได้"):
                st.write(all_errors)

        st.subheader("ดาวน์โหลดผลลัพธ์")
        st.caption("เลือกชนิดไฟล์ตามลักษณะงานของคุณ ข้อมูลรายงานหลักเหมือนกันทั้งสองแบบ")
        download_excel, download_csv = st.columns(2, gap="large")
        with download_excel:
            st.markdown('<div class="download-heading">Excel (.xlsx)</div>', unsafe_allow_html=True)
            st.markdown(
                '<div class="download-detail">เหมาะสำหรับเปิด ตรวจสอบ กรอง และจัดรูปแบบต่อใน Excel<br>วางเลขใบอนุญาตในคอลัมน์ W เพื่อดูผลในคอลัมน์ X</div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                label="ดาวน์โหลด Excel",
                data=excel_data,
                file_name="NBTC_Report_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="download_excel",
            )
        with download_csv:
            st.markdown('<div class="download-heading">CSV (.csv)</div>', unsafe_allow_html=True)
            st.markdown(
                '<div class="download-detail">เหมาะสำหรับนำเข้าโปรแกรมหรือฐานข้อมูลอย่างรวดเร็ว<br>มีเฉพาะรายงานหลัก 22 คอลัมน์</div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                label="ดาวน์โหลด CSV",
                data=report_csv,
                file_name="NBTC_Report_Summary.csv",
                mime="text/csv",
                key="download_csv",
            )
