import streamlit as st

st.set_page_config(page_title="กสทช. PDF to CSV", page_icon="📶", layout="wide")

try:
    from extractor import (
        QA_COLUMNS,
        REPORT_COLUMNS,
        build_qa_records,
        build_report_records,
        create_csv_bytes,
        parse_pdf_bytes,
    )
except ModuleNotFoundError as error:
    if error.name == "extractor":
        st.error("ไม่พบไฟล์ extractor.py ในโปรเจกต์ที่นำขึ้น Streamlit Cloud")
        st.info("เพิ่มไฟล์ extractor.py ไว้โฟลเดอร์เดียวกับ app.py แล้ว commit และ redeploy อีกครั้ง")
    elif error.name == "pdfplumber":
        st.error("ยังไม่ได้ติดตั้งไลบรารี pdfplumber")
        st.info("ตรวจว่ามี requirements.txt อยู่โฟลเดอร์เดียวกับ app.py และมีบรรทัด pdfplumber>=0.11,<1")
    else:
        st.error(f"ไม่พบไลบรารีที่จำเป็น: {error.name}")
    st.stop()
except ImportError:
    st.error("ไฟล์ extractor.py ใน Streamlit Cloud ไม่ตรงกับ app.py เวอร์ชันปัจจุบัน")
    st.info("commit app.py, extractor.py และ requirements.txt เวอร์ชันล่าสุด แล้ว redeploy อีกครั้ง")
    st.stop()

st.title("ระบบสกัดรายงาน กสทช. เป็น CSV")
st.write(
    "อัปโหลดรายงาน PDF ได้หลายไฟล์ ระบบจะแยกข้อมูลทีละหน้า "
    "จึงรองรับทั้งรายงาน 1 สถานีต่อไฟล์และหลายสถานีในไฟล์เดียว โดยส่งออกเป็น CSV"
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
    "รูปแบบข้อมูลใน CSV",
    options=("สรุป 1 แถวต่อสถานี", "ละเอียด 1 แถวต่อความถี่"),
    horizontal=True,
)

if st.button("เริ่มสกัดข้อมูล", type="primary", disabled=not uploaded_files):
    all_stations = []
    all_errors = []

    with st.spinner("กำลังอ่านและจัดรูปแบบรายงาน..."):
        for uploaded_file in uploaded_files:
            stations, errors = parse_pdf_bytes(uploaded_file.name, uploaded_file.getvalue())
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
        qa_csv = create_csv_bytes(qa_records, QA_COLUMNS)

        needs_review = sum(1 for row in qa_records if row["สถานะ"] == "ต้องตรวจสอบ")
        left, middle, right = st.columns(3)
        left.metric("สถานีที่พบ", len(all_stations))
        middle.metric("แถวใน CSV", len(report_records))
        right.metric("หน้าที่ต้องตรวจสอบ", needs_review)

        st.subheader("ตัวอย่างข้อมูล")
        st.dataframe(report_records[:20], use_container_width=True, hide_index=True)

        if needs_review:
            st.warning("บางหน้าอ่านข้อมูลได้ไม่ครบ โปรดตรวจไฟล์ผลตรวจสอบก่อนนำรายงานไปใช้งาน")
            st.dataframe(
                [row for row in qa_records if row["สถานะ"] == "ต้องตรวจสอบ"],
                use_container_width=True,
                hide_index=True,
            )

        if all_errors:
            with st.expander(f"รายละเอียด {len(all_errors)} หน้าที่อ่านไม่ได้"):
                st.write(all_errors)

        download_report, download_qa = st.columns(2)
        with download_report:
            st.download_button(
                label="ดาวน์โหลดรายงาน CSV",
                data=report_csv,
                file_name="NBTC_Report_Summary.csv",
                mime="text/csv",
                type="primary",
            )
        with download_qa:
            st.download_button(
                label="ดาวน์โหลดผลตรวจสอบ CSV",
                data=qa_csv,
                file_name="NBTC_Report_Validation.csv",
                mime="text/csv",
            )
