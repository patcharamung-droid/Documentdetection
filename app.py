import streamlit as st

from extractor import (
    build_qa_records,
    build_report_records,
    create_excel_bytes,
    parse_pdf_bytes,
)


st.set_page_config(page_title="กสทช. PDF to CSV", page_icon="📶", layout="wide")

st.title("ระบบสกัดรายงาน กสทช. เป็น Excel")
st.write(
    "อัปโหลดรายงาน PDF ได้หลายไฟล์ ระบบจะแยกข้อมูลทีละหน้า "
    "จึงรองรับทั้งรายงาน 1 สถานีต่อไฟล์และหลายสถานีในไฟล์เดียว"
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
    "รูปแบบข้อมูลใน Excel",
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
        excel_data = create_excel_bytes(report_records, qa_records)

        needs_review = sum(1 for row in qa_records if row["สถานะ"] == "ต้องตรวจสอบ")
        left, middle, right = st.columns(3)
        left.metric("สถานีที่พบ", len(all_stations))
        middle.metric("แถวใน Excel", len(report_records))
        right.metric("หน้าที่ต้องตรวจสอบ", needs_review)

        st.subheader("ตัวอย่างข้อมูล")
        st.dataframe(report_records[:20], use_container_width=True, hide_index=True)

        if needs_review:
            st.warning("บางหน้าอ่านข้อมูลได้ไม่ครบ โปรดเปิดแผ่นงาน “ตรวจสอบข้อมูล” ใน Excel ก่อนนำไปใช้งาน")
            st.dataframe(
                [row for row in qa_records if row["สถานะ"] == "ต้องตรวจสอบ"],
                use_container_width=True,
                hide_index=True,
            )

        if all_errors:
            with st.expander(f"รายละเอียด {len(all_errors)} หน้าที่อ่านไม่ได้"):
                st.write(all_errors)

        st.download_button(
            label="ดาวน์โหลดไฟล์ Excel",
            data=excel_data,
            file_name="NBTC_Report_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
