import streamlit as st
import pandas as pd
import openpyxl
import io

# ---------------------------------------------------------
# 1. ตั้งค่าหน้าเว็บและสไตล์ CSS (Page Configuration & Custom CSS)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Retail Audit & Survey Suite",
    page_icon="🏬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling ให้หน้าตาดูโมเดิร์น สบายตา รองรับการใช้งาน
st.markdown("""
<style>
    /* แบ็กกราวด์หลัก */
    .stApp {
        background-color: #f8fafc;
        font-family: 'Sarabun', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Header การ์ดเมนู Sidebar */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    
    /* กล่องข้อความการ์ด */
    .custom-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        border: 1px solid #e2e8f0;
        margin-bottom: 20px;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .custom-card:hover {
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08), 0 4px 6px -2px rgba(0, 0, 0, 0.04);
    }
    
    /* สไตล์ระยะทาง GPS */
    .distance-alert {
        color: #ef4444;
        font-weight: 700;
        background-color: #fef2f2;
        padding: 2px 8px;
        border-radius: 6px;
    }
    .distance-ok {
        color: #10b981;
        font-weight: 700;
        background-color: #ecfdf5;
        padding: 2px 8px;
        border-radius: 6px;
    }
    
    /* หัวข้อและ Tag */
    .badge-tag {
        background-color: #e0f2fe;
        color: #0369a1;
        font-weight: 600;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# 2. ฟังก์ชันช่วยประมวลผลข้อมูล (Helper Functions)
# ---------------------------------------------------------

# ถอด @st.cache_data ออกเพื่อรองรับ Multi-user ป้องกันข้อมูลค้างข้ามผู้ใช้
def load_spirits_data(file):
    filename = file.name.lower()
    if filename.endswith('.csv'):
        try:
            df = pd.read_csv(file, header=1)
        except Exception:
            file.seek(0)
            df = pd.read_csv(file, encoding='tis-620', header=1)
    elif filename.endswith(('.xlsx', '.xlsm')):
        df = pd.read_excel(file, engine='openpyxl', header=1)
    else:
        return None
        
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df = df.dropna(how='all')
    return df


# ถอด @st.cache_data ออกเพื่อรองรับ Multi-user ป้องกันข้อมูลค้างข้ามผู้ใช้
def load_checkin_data(file_bytes):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
    
    header_row_idx = 4
    for idx, row in df_raw.iterrows():
        row_str_values = [str(val) for val in row.values]
        if 'รูปถ่าย Check-In' in row_str_values or 'ลำดับที่' in row_str_values:
            header_row_idx = idx
            break
            
    headers = df_raw.iloc[header_row_idx].values
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, skiprows=header_row_idx + 1, header=None)
    df.columns = [str(h).strip() if pd.notna(h) else f"Col_{i}" for i, h in enumerate(headers[:df.shape[1]])]
    df = df.dropna(how='all').reset_index(drop=True)
    
    # ดึง Hyperlinks จาก OpenPyXL
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    sheet = wb.active
    
    col_map = {name: idx + 1 for idx, name in enumerate(df.columns)}
    checkin_col_idx = col_map.get('รูปถ่าย Check-In')
    closed_col_idx = col_map.get('รูปสถานะร้านปิด/เลิกกิจการ')
    
    checkin_urls = []
    closed_urls = []
    data_start_row = header_row_idx + 2
    
    for r_idx in range(len(df)):
        excel_row = data_start_row + r_idx
        
        # รูปถ่าย Check-In
        ch_url = None
        if checkin_col_idx:
            cell = sheet.cell(row=excel_row, column=checkin_col_idx)
            if cell.hyperlink and cell.hyperlink.target:
                ch_url = cell.hyperlink.target
            elif isinstance(cell.value, str) and cell.value.startswith(('http://', 'https://')):
                ch_url = cell.value
        checkin_urls.append(ch_url if ch_url else df.iloc[r_idx].get('รูปถ่าย Check-In', None))
        
        # รูปสถานะร้านปิด
        cl_url = None
        if closed_col_idx:
            cell = sheet.cell(row=excel_row, column=closed_col_idx)
            if cell.hyperlink and cell.hyperlink.target:
                cl_url = cell.hyperlink.target
            elif isinstance(cell.value, str) and cell.value.startswith(('http://', 'https://')):
                cl_url = cell.value
        closed_urls.append(cl_url if cl_url else df.iloc[r_idx].get('รูปสถานะร้านปิด/เลิกกิจการ', None))
        
    df['URL_CheckIn_Extracted'] = checkin_urls
    df['URL_Closed_Extracted'] = closed_urls
    return df


# ---------------------------------------------------------
# 3. เมนูหลัก Navigation (Sidebar)
# ---------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/shop.png", width=64)
    st.title("R4 Management & Analytics")
    st.caption("ระบบช่วยบริหารจัดการตรวจสอบรายงาน สถิติ R4")
    st.markdown("---")
    
    menu = st.radio(
        "📌 เลือกเมนูใช้งาน:",
        ["🍹 Spirits _ สำรวจการติดสื่อในร้านค้า", "🔍 CheckIn Audit Pro รายงานเยี่ยมร้านค้า"],
        index=0
    )
    st.markdown("---")


# =========================================================
# 4. เมนูที่ 1: Spirits _ สำรวจการติดสื่อในร้านค้า
# =========================================================
if menu == "🍹 Spirits _ สำรวจการติดสื่อในร้านค้า":
    st.title("🍹 Spirits — สำรวจการติดสื่อในร้านค้า")
    st.caption("ตรวจสอบรูปถ่ายการจัดตั้งสื่อโฆษณาตามแบรนด์สินค้าในแต่ละร้านค้า")

    uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์สำรวจ (.xlsm, .xlsx, .csv)", type=["csv", "xlsx", "xlsm"], key="spirits_upload")

    if uploaded_file is not None:
        df = load_spirits_data(uploaded_file)
        
        if df is not None:
            # จัดการชื่อพนักงาน
            if 'ชื่อ' in df.columns and 'นามสกุล' in df.columns:
                df['พนักงาน'] = df['ชื่อ'].fillna('').astype(str) + ' ' + df['นามสกุล'].fillna('').astype(str)
                df['พนักงาน'] = df['พนักงาน'].str.strip().replace('', 'ไม่ระบุ')
            elif 'createdBy' in df.columns:
                df['พนักงาน'] = df['createdBy'].fillna('ไม่ระบุ')
            else:
                df['พนักงาน'] = "ไม่ระบุ"

            # Sidebar Filters
            st.sidebar.header("🎯 ตัวกรองข้อมูล")
            
            staff_list = ["ทั้งหมด"] + sorted([x for x in df['พนักงาน'].unique() if x != 'ไม่ระบุ'])
            selected_staff = st.sidebar.selectbox("พนักงานผู้สำรวจ:", staff_list)

            if 'shopName' in df.columns:
                shop_list = ["ทั้งหมด"] + sorted(list(df['shopName'].dropna().astype(str).unique()))
                selected_shop = st.sidebar.selectbox("เลือกร้านค้า:", shop_list)
            else:
                selected_shop = "ทั้งหมด"

            if 'region' in df.columns:
                region_list = ["ทั้งหมด"] + sorted(list(df['region'].dropna().astype(str).unique()))
                selected_region = st.sidebar.selectbox("Region:", region_list)
            else:
                selected_region = "ทั้งหมด"

            # Filter Logic
            filtered_df = df.copy()
            if selected_staff != "ทั้งหมด":
                filtered_df = filtered_df[filtered_df['พนักงาน'] == selected_staff]
            if selected_shop != "ทั้งหมด":
                filtered_df = filtered_df[filtered_df['shopName'].astype(str) == selected_shop]
            if selected_region != "ทั้งหมด":
                filtered_df = filtered_df[filtered_df['region'].astype(str) == selected_region]

            # KPI Dashboard Cards
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("จำนวนรายการสำรวจ", f"{len(filtered_df):,} รายการ")
            kpi2.metric("ร้านค้าที่สำรวจ", f"{filtered_df['shopCode'].nunique() if 'shopCode' in filtered_df.columns else 0:,} ร้าน")
            kpi3.metric("พนักงานผู้สำรวจ", f"{filtered_df['พนักงาน'].nunique():,} คน")

            st.markdown("---")

            # แสดงผลการ์ดร้านค้าและรูปภาพ
            st.subheader("📸 รายการสำรวจสื่อโฆษณา")
            
            if filtered_df.empty:
                st.warning("⚠️ ไม่พบข้อมูลตามเงื่อนไขที่เลือก")
            else:
                img_cols = [c for c in filtered_df.columns if 'ถ่ายรูปติดสื่อ' in str(c)]

                for idx, row in filtered_df.iterrows():
                    shop_name_val = row.get('shopName', 'ไม่ระบุร้านค้า')
                    shop_code_val = row.get('shopCode', '-')
                    staff_val = row.get('พนักงาน', '-')

                    with st.expander(f"📍 **{shop_name_val}** (รหัสร้าน: {shop_code_val}) | พนักงาน: {staff_val}"):
                        st.markdown(f"""
                        * **Transaction ID:** `{row.get('transactionId', '-')}`
                        * **วันที่สำรวจ:** {row.get('createdDate', '-')}
                        * **BU / Channel / Region:** {row.get('BU', '-')} / {row.get('channel', '-')} / {row.get('region', '-')}
                        * **พิกัด GPS:** Lat {row.get('lat', '-')}, Long {row.get('long', '-')}
                        """)
                        
                        st.markdown("##### 🖼️ รูปถ่ายสื่อการขายในร้านค้า")
                        has_image = False
                        img_grid = st.columns(3)
                        col_idx = 0
                        
                        for col_name in img_cols:
                            img_url = row[col_name]
                            if pd.notna(img_url) and isinstance(img_url, str) and img_url.strip().startswith(('http://', 'https://')):
                                has_image = True
                                with img_grid[col_idx % 3]:
                                    st.image(img_url.strip(), caption=col_name, use_container_width=True)
                                col_idx += 1

                        if not has_image:
                            st.info("ℹ️ ไม่มีรูปภาพติดสื่อถูกส่งมาในรายการนี้")

                with st.expander("📊 ดูตารางข้อมูลดิบ (Raw Data)"):
                    st.dataframe(filtered_df, use_container_width=True)
    else:
        st.info("👆 กรุณาอัปโหลดไฟล์รายงานการสำรวจสื่อ (.xlsm, .xlsx, .csv) ที่เมนูด้านซ้ายเพื่อเริ่มใช้งาน")


# =========================================================
# 5. เมนูที่ 2: CheckIn Audit Pro
# =========================================================
elif menu == "🔍 CheckIn Audit Pro รายงานเยี่ยมร้านค้า":
    st.title("🔍 CheckIn Audit Pro รายงานเยี่ยมร้านค้า")
    st.caption("ระบบตรวจสอบรูปถ่าย Check-In ระยะห่างพิกัด GPS และสถานะร้านค้า")

    uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์รายงานการเยี่ยมร้านค้า (.xlsx)", type=["xlsx"], key="checkin_upload")

    if uploaded_file is not None:
        # การป้องกัน Multi-user ปัญหาข้อมูลตีกัน: สร้าง Unique File Identifier
        file_identifier = f"{uploaded_file.name}_{uploaded_file.size}"
        
        # ถ้าระบบตรวจพบว่าผู้ใช้เปลี่ยนไฟล์ หรือโหลดไฟล์ใหม่ ให้รีเซ็ต Session Audit ผลการตรวจทันที
        if 'current_checkin_file_id' not in st.session_state or st.session_state['current_checkin_file_id'] != file_identifier:
            st.session_state['current_checkin_file_id'] = file_identifier
            st.session_state['audit_results'] = {}

        try:
            file_bytes = uploaded_file.read()
            df = load_checkin_data(file_bytes)

            st.sidebar.header("🎯 ตัวกรองการตรวจสอบ")
            search_query = st.sidebar.text_input("🔍 ค้นหา (ชื่อร้าน / รหัสร้าน / พนักงาน)", "")
            
            filter_distance = st.sidebar.checkbox("กรองเฉพาะที่เข้านอกพิกัด")
            dist_threshold = st.sidebar.number_input("ระยะทางเกินกำหนด (เมตร)", value=100, step=50) if filter_distance else 0
            
            status_options = ["ทั้งหมด"] + list(df['สถานะร้านปิด/เลิกกิจการ'].dropna().unique()) if 'สถานะร้านปิด/เลิกกิจการ' in df.columns else ["ทั้งหมด"]
            selected_status = st.sidebar.selectbox("สถานะร้านค้า", status_options)
            
            audit_filter = st.sidebar.selectbox("สถานะการตรวจ", ["ทั้งหมด", "ยังไม่ได้ตรวจ", "ผ่าน", "ไม่ผ่าน", "รอตรวจสอบเพิ่ม"])
            
            st.sidebar.markdown("---")
            items_per_page = st.sidebar.selectbox("แสดงรายการต่อหน้า", [12, 24, 48, 96], index=0)

            # Filtering
            filtered_df = df.copy()
            if search_query:
                mask = filtered_df.astype(str).apply(lambda row: row.str.contains(search_query, case=False, na=False)).any(axis=1)
                filtered_df = filtered_df[mask]
                
            if filter_distance and 'ระยะห่างเข้า (เมตร)' in filtered_df.columns:
                filtered_df['ระยะห่างเข้า (เมตร)'] = pd.to_numeric(filtered_df['ระยะห่างเข้า (เมตร)'], errors='coerce').fillna(0)
                filtered_df = filtered_df[filtered_df['ระยะห่างเข้า (เมตร)'] >= dist_threshold]
                
            if selected_status != "ทั้งหมด" and 'สถานะร้านปิด/เลิกกิจการ' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['สถานะร้านปิด/เลิกกิจการ'] == selected_status]

            def get_audit_status(index):
                return st.session_state['audit_results'].get(index, {}).get('status', 'ยังไม่ได้ตรวจ')

            if audit_filter != "ทั้งหมด":
                filtered_df['Current_Audit_Status'] = filtered_df.index.map(get_audit_status)
                filtered_df = filtered_df[filtered_df['Current_Audit_Status'] == audit_filter]

            # KPI Summary
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("จำนวนทั้งหมด", f"{len(df):,} รายการ")
            col_m2.metric("ตรงตามตัวกรอง", f"{len(filtered_df):,} รายการ")
            audited_count = len([v for v in st.session_state['audit_results'].values() if v.get('status') != 'ยังไม่ได้ตรวจ'])
            col_m3.metric("ตรวจแล้ว", f"{audited_count:,} รายการ")
            col_m4.metric("คงเหลือ", f"{max(0, len(df) - audited_count):,} รายการ")

            st.markdown("---")

            # Pagination
            total_items = len(filtered_df)
            total_pages = max(1, (total_items + items_per_page - 1) // items_per_page)
            
            col_p1, col_p2 = st.columns([1, 4])
            with col_p1:
                current_page = st.number_input(f"หน้า (จากทั้งหมด {total_pages})", min_value=1, max_value=total_pages, value=1)
            
            start_idx = (current_page - 1) * items_per_page
            end_idx = start_idx + items_per_page
            page_data = filtered_df.iloc[start_idx:end_idx]

            # Grid Cards Display
            cols_per_row = 3
            grid_cols = st.columns(cols_per_row)

            for i, (orig_idx, row) in enumerate(page_data.iterrows()):
                with grid_cols[i % cols_per_row]:
                    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
                    
                    shop_name = row.get('ชื่อร้านค้า', 'ไม่ระบุชื่อร้าน')
                    shop_code = row.get('รหัสร้านค้า', '-')
                    staff_name = row.get('ชื่อพนักงาน', '-')
                    visit_date = row.get('วันที่เข้าเยี่ยม', '-')
                    visit_time = row.get('เวลาเข้า', '-')
                    distance = row.get('ระยะห่างเข้า (เมตร)', 0)
                    
                    st.subheader(f"🏪 {shop_name}")
                    st.caption(f"รหัสร้าน: {shop_code} | พนักงาน: {staff_name}")
                    st.write(f"📅 **วันที่:** {visit_date} | ⏰ **เวลา:** {visit_time}")
                    
                    # GPS Distance Check
                    try:
                        dist_val = float(distance)
                        if dist_val > 100:
                            st.markdown(f"📍 ระยะห่าง: <span class='distance-alert'>{dist_val:.1f} เมตร (เกินเกณฑ์)</span>", unsafe_allow_html=True)
                            reason = row.get('เหตุผลการเข้านอกพิกัด', '-')
                            if pd.notna(reason) and str(reason).strip() != '-':
                                st.caption(f"เหตุผลเข้านอกพิกัด: {reason}")
                        else:
                            st.markdown(f"📍 ระยะห่าง: <span class='distance-ok'>{dist_val:.1f} เมตร</span>", unsafe_allow_html=True)
                    except Exception:
                        st.write(f"📍 ระยะห่าง: {distance}")

                    # Images
                    checkin_img = row.get('URL_CheckIn_Extracted')
                    closed_img = row.get('URL_Closed_Extracted')
                    
                    st.markdown("#### 📷 รูปถ่าย Check-In")
                    if pd.notna(checkin_img) and str(checkin_img).startswith("http"):
                        st.image(str(checkin_img), use_container_width=True)
                    else:
                        st.warning("⚠️ ไม่มี URL รูปถ่าย Check-In")
                    
                    if pd.notna(closed_img) and str(closed_img).startswith("http"):
                        with st.expander("🖼️ ดูรูปสถานะร้านปิด/เลิกกิจการ"):
                            st.image(str(closed_img), use_container_width=True)

                    # Audit Action Form
                    current_result = st.session_state['audit_results'].get(orig_idx, {'status': 'ยังไม่ได้ตรวจ', 'note': ''})
                    
                    options = ["ยังไม่ได้ตรวจ", "ผ่าน", "ไม่ผ่าน", "รอตรวจสอบเพิ่ม"]
                    default_index = options.index(current_result['status']) if current_result['status'] in options else 0

                    # ผูก key ด้วย file_identifier เพื่อกันความผิดพลาดข้ามการอัปโหลด
                    status_choice = st.radio(
                        "ผลการตรวจสอบ:",
                        options,
                        index=default_index,
                        key=f"status_{file_identifier}_{orig_idx}"
                    )
                    
                    audit_note = st.text_input(
                        "หมายเหตุเพิ่มเติม:",
                        value=current_result['note'],
                        key=f"note_{file_identifier}_{orig_idx}",
                        placeholder="เช่น ถ่ายรูปไม่เห็นหน้าร้าน"
                    )
                    
                    st.session_state['audit_results'][orig_idx] = {
                        'status': status_choice,
                        'note': audit_note
                    }
                    
                    st.markdown('</div>', unsafe_allow_html=True)

            # Export Button
            st.markdown("---")
            st.subheader("📥 ส่งออกผลการตรวจสอบ (Export Results)")
            
            if st.button("🔄 ประมวลผลและเตรียมไฟล์ดาวน์โหลด"):
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
