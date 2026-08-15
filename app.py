import streamlit as st
import pandas as pd
import openpyxl
import io

# ---------------------------------------------------------
# ตั้งค่าหน้าเว็บและสไตล์ (Page Configuration & Styling)
# ---------------------------------------------------------
st.set_page_config(
    page_title="ระบบตรวจสอบรูปถ่าย Check-In ร้านค้า",
    page_icon="🔍",
    layout="wide"
)

st.markdown("""
<style>
    .stApp {
        background-color: #f8fafc;
    }
    .card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 16px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06);
        border: 1px solid #e2e8f0;
        margin-bottom: 20px;
    }
    .distance-alert {
        color: #dc2626;
        font-weight: bold;
    }
    .distance-ok {
        color: #16a34a;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# ฟังก์ชันช่วยประมวลผลไฟล์ Excel (Helper Functions)
# ---------------------------------------------------------
def load_excel_data(uploaded_file):
    """
    อ่านไฟล์ Excel ดึงข้อมูลแถวหัวข้อ (Header) อัตโนมัติ
    และดึงลิงก์รูปภาพ (Hyperlink/URL) ทั้งหมดออกมา
    """
    file_bytes = uploaded_file.read()
    
    # 1. อ่านข้อมูลเพื่อค้นหาแถวที่เป็นหัวข้อตาราง (Header)
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
    
    header_row_idx = 4  # ค่าน่าเริ่มต้นคือแถวที่ 5 ใน Excel (Index = 4)
    for idx, row in df_raw.iterrows():
        row_str_values = [str(val) for val in row.values]
        if 'รูปถ่าย Check-In' in row_str_values or 'ลำดับที่' in row_str_values:
            header_row_idx = idx
            break
            
    # 2. ดึงข้อมูลจริงโดยข้ามแถวส่วนหัว
    headers = df_raw.iloc[header_row_idx].values
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, skiprows=header_row_idx + 1, header=None)
    df.columns = headers[:df.shape[1]]
    
    # ลบแถวที่เป็นค่าว่างทั้งหมด
    df = df.dropna(how='all').reset_index(drop=True)
    
    # 3. ดึง Hyperlinks จาก OpenPyXL (กรณีลิงก์ฝังในข้อความ)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    sheet = wb.active
    
    col_map = {name: idx + 1 for idx, name in enumerate(df.columns)}
    
    checkin_col_idx = col_map.get('รูปถ่าย Check-In')
    closed_col_idx = col_map.get('รูปสถานะร้านปิด/เลิกกิจการ')
    
    checkin_urls = []
    closed_urls = []
    
    data_start_row = header_row_idx + 2  # OpenPyXL เริ่มนับจากแถวที่ 1
    
    for r_idx in range(len(df)):
        excel_row = data_start_row + r_idx
        
        # รูปถ่าย Check-In
        ch_url = None
        if checkin_col_idx:
            cell = sheet.cell(row=excel_row, column=checkin_col_idx)
            if cell.hyperlink:
                ch_url = cell.hyperlink.target
            elif isinstance(cell.value, str) and cell.value.startswith('http'):
                ch_url = cell.value
        checkin_urls.append(ch_url if ch_url else df.iloc[r_idx].get('รูปถ่าย Check-In', None))
        
        # รูปสถานะร้านปิด/เลิกกิจการ
        cl_url = None
        if closed_col_idx:
            cell = sheet.cell(row=excel_row, column=closed_col_idx)
            if cell.hyperlink:
                cl_url = cell.hyperlink.target
            elif isinstance(cell.value, str) and cell.value.startswith('http'):
                cl_url = cell.value
        closed_urls.append(cl_url if cl_url else df.iloc[r_idx].get('รูปสถานะร้านปิด/เลิกกิจการ', None))
        
    df['URL_CheckIn_Extracted'] = checkin_urls
    df['URL_Closed_Extracted'] = closed_urls
    
    return df


# ---------------------------------------------------------
# ส่วนแสดงผลหน้าเว็บ (App Interface)
# ---------------------------------------------------------
st.title("🔍 ระบบตรวจสอบรูปถ่าย Check-In ร้านค้า")
st.caption("เครื่องมือสำหรับตรวจเช็กความถูกต้องของรูปถ่ายการเยี่ยมร้านค้า พิกัด GPS และสถานะร้าน")

# บันทึกผลการตรวจลงใน Session State
if 'audit_results' not in st.session_state:
    st.session_state['audit_results'] = {}

# แถบเมนูด้านข้าง (Sidebar)
with st.sidebar:
    st.header("📁 อัปโหลดไฟล์ Excel")
    uploaded_file = st.file_uploader("เลือกไฟล์รายงานการเยี่ยมร้านค้า (.xlsx)", type=["xlsx"])
    
    st.markdown("---")
    st.header("🎯 ตัวกรองการตรวจสอบ")

if uploaded_file is not None:
    try:
        df = load_excel_data(uploaded_file)
        
        with st.sidebar:
            # ช่องค้นหา
            search_query = st.text_input("🔍 ค้นหา (ชื่อร้าน / รหัสร้าน / พนักงาน)", "")
            
            # ตัวกรองระยะทาง
            filter_distance = st.checkbox("กรองเฉพาะที่เข้านอกพิกัด (ระยะทางเกินกำหนด)")
            dist_threshold = st.number_input("ระยะทางที่ถือว่าเกินพิกัด (เมตร)", value=100, step=50) if filter_distance else 0
            
            # ตัวกรองสถานะร้านค้า
            status_options = ["ทั้งหมด"] + list(df['สถานะร้านปิด/เลิกกิจการ'].dropna().unique()) if 'สถานะร้านปิด/เลิกกิจการ' in df.columns else ["ทั้งหมด"]
            selected_status = st.selectbox("สถานะร้านค้า", status_options)
            
            # ตัวกรองสถานะการตรวจ
            audit_filter = st.selectbox("สถานะการตรวจ", ["ทั้งหมด", "ยังไม่ได้ตรวจ", "ผ่าน", "ไม่ผ่าน", "รอตรวจสอบเพิ่ม"])
            
            # การแบ่งหน้า (Pagination)
            st.markdown("---")
            items_per_page = st.selectbox("แสดงผลกี่รายการต่อหน้า", [12, 24, 48, 96], index=0)

        # ---------------------------------------------------------
        # ระบบการกรองข้อมูล (Filtering Logic)
        # ---------------------------------------------------------
        filtered_df = df.copy()
        
        if search_query:
            mask = filtered_df.astype(str).apply(lambda row: row.str.contains(search_query, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]
            
        if filter_distance and 'ระยะห่างเข้า (เมตร)' in filtered_df.columns:
            filtered_df['ระยะห่างเข้า (เมตร)'] = pd.to_numeric(filtered_df['ระยะห่างเข้า (เมตร)'], errors='coerce').fillna(0)
            filtered_df = filtered_df[filtered_df['ระยะห่างเข้า (เมตร)'] >= dist_threshold]
            
        if selected_status != "ทั้งหมด" and 'สถานะร้านปิด/เลิกกิจการ' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['สถานะร้านปิด/เลิกกิจการ'] == selected_status]

        # กรองตามสถานะการตรวจ
        def get_audit_status(index):
            return st.session_state['audit_results'].get(index, {}).get('status', 'ยังไม่ได้ตรวจ')

        if audit_filter != "ทั้งหมด":
            filtered_df['Current_Audit_Status'] = filtered_df.index.map(get_audit_status)
            filtered_df = filtered_df[filtered_df['Current_Audit_Status'] == audit_filter]

        # แสดงสรุปตัวเลข (Metrics)
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("จำนวนรายการทั้งหมด", f"{len(df)} รายการ")
        col_m2.metric("ตรงตามเงื่อนไขค้นหา", f"{len(filtered_df)} รายการ")
        
        audited_count = len(st.session_state['audit_results'])
        col_m3.metric("ตรวจแล้ว", f"{audited_count} รายการ")
        col_m4.metric("คงเหลือ", f"{len(df) - audited_count} รายการ")

        st.markdown("---")

        # ---------------------------------------------------------
        # การจัดหน้ารายการ (Pagination)
        # ---------------------------------------------------------
        total_items = len(filtered_df)
        total_pages = max(1, (total_items + items_per_page - 1) // items_per_page)
        
        col_p1, col_p2 = st.columns([1, 4])
        with col_p1:
            current_page = st.number_input(f"หน้า (จากทั้งหมด {total_pages})", min_value=1, max_value=total_pages, value=1)
        
        start_idx = (current_page - 1) * items_per_page
        end_idx = start_idx + items_per_page
        page_data = filtered_df.iloc[start_idx:end_idx]

        # ---------------------------------------------------------
        # แสดงรายการแบบการ์ด (Grid Display)
        # ---------------------------------------------------------
        cols_per_row = 3
        grid_cols = st.columns(cols_per_row)

        for i, (orig_idx, row) in enumerate(page_data.iterrows()):
            with grid_cols[i % cols_per_row]:
                st.markdown('<div class="card">', unsafe_allow_html=True)
                
                # ข้อมูลหัวข้อการ์ด
                shop_name = row.get('ชื่อร้านค้า', 'ไม่ระบุชื่อร้าน')
                shop_code = row.get('รหัสร้านค้า', '-')
                staff_name = row.get('ชื่อพนักงาน', '-')
                visit_date = row.get('วันที่เข้าเยี่ยม', '-')
                visit_time = row.get('เวลาเข้า', '-')
                distance = row.get('ระยะห่างเข้า (เมตร)', 0)
                
                st.subheader(f"🏪 {shop_name}")
                st.caption(f"รหัสร้าน: {shop_code} | พนักงาน: {staff_name}")
                st.write(f"📅 **วันที่:** {visit_date} | ⏰ **เวลา:** {visit_time}")
                
                # แสดงสถานะระยะห่าง GPS
                try:
                    dist_val = float(distance)
                    if dist_val > 100:
                        st.markdown(f"📍 ระยะห่าง: <span class='distance-alert'>{dist_val:.1f} เมตร (เกินเกณฑ์)</span>", unsafe_allow_html=True)
                        reason = row.get('เหตุผลการเข้านอกพิกัด', '-')
                        if pd.notna(reason) and reason != '-':
                            st.caption(f"เหตุผลเข้านอกพิกัด: {reason}")
                    else:
                        st.markdown(f"📍 ระยะห่าง: <span class='distance-ok'>{dist_val:.1f} เมตร</span>", unsafe_allow_html=True)
                except Exception:
                    st.write(f"📍 ระยะห่าง: {distance}")

                # แสดงรูปถ่าย Check-In
                checkin_img = row.get('URL_CheckIn_Extracted')
                closed_img = row.get('URL_Closed_Extracted')
                
                st.markdown("#### 📷 รูปถ่าย Check-In")
                if pd.notna(checkin_img) and str(checkin_img).startswith("http"):
                    st.image(str(checkin_img), use_container_width=True)
                else:
                    st.warning("⚠️ ไม่มี URL รูปถ่าย Check-In หรือรูปภาพไม่สมบูรณ์")
                
                # แสดงรูปสถานะร้านปิด (ถ้ามี)
                if pd.notna(closed_img) and str(closed_img).startswith("http"):
                    with st.expander("🖼️ ดูรูปสถานะร้านปิด/เลิกกิจการ"):
                        st.image(str(closed_img), use_container_width=True)

                # แบบฟอร์มลงผลการตรวจ
                current_result = st.session_state['audit_results'].get(orig_idx, {'status': 'ยังไม่ได้ตรวจ', 'note': ''})
                
                status_choice = st.radio(
                    "ผลการตรวจสอบ:",
                    ["ยังไม่ได้ตรวจ", "ผ่าน", "ไม่ผ่าน", "รอตรวจสอบเพิ่ม"],
                    index=["ยังไม่ได้ตรวจ", "ผ่าน", "ไม่ผ่าน", "รอตรวจสอบเพิ่ม"].index(current_result['status']),
                    key=f"status_{orig_idx}"
                )
                
                audit_note = st.text_input(
                    "หมายเหตุเพิ่มเติม:",
                    value=current_result['note'],
                    key=f"note_{orig_idx}",
                    placeholder="เช่น ถ่ายรูปไม่เห็นหน้าร้าน / รูปดำ"
                )
                
                # บันทึกข้อมูลลงใน Session State
                st.session_state['audit_results'][orig_idx] = {
                    'status': status_choice,
                    'note': audit_note
                }
                
                st.markdown('</div>', unsafe_allow_html=True)

        # ---------------------------------------------------------
        # ดาวน์โหลดผลการตรวจสอบ (Export Results)
        # ---------------------------------------------------------
        st.markdown("---")
        st.header("📥 ส่งออกผลการตรวจสอบ (Export Results)")
        
        if st.button("ประมวลผลและเตรียมไฟล์ดาวน์โหลด"):
            export_df = df.copy()
            export_df['ผลการตรวจ'] = export_df.index.map(lambda x: st.session_state['audit_results'].get(x, {}).get('status', 'ยังไม่ได้ตรวจ'))
            export_df['หมายเหตุการตรวจ'] = export_df.index.map(lambda x: st.session_state['audit_results'].get(x, {}).get('note', ''))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Audit_Results')
            
            st.download_button(
                label="⬇️ ดาวน์โหลดรายงานผลการตรวจ (.xlsx)",
                data=output.getvalue(),
                file_name="รายงานการตรวจรูปถ่าย_CheckIn.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
else:
    st.info("👋 กรุณาอัปโหลดไฟล์ Excel รายงานการเยี่ยมร้านค้าที่แถบด้านซ้ายเพื่อเริ่มต้นใช้งาน")