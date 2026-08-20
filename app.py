import streamlit as st
import pandas as pd
import os
import base64
from streamlit_option_menu import option_menu

# ==============================
# STEP 1: GLOBAL CONFIG & CUSTOM CSS (GEOMETRIC & TANK BG)
# ==============================
st.set_page_config(
    page_title="HEIL TRAILER | Executive Inventory Control",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.cache_data.clear()

# ฟังก์ชั่นแปลงรูปเป็น Base64 แบบรองรับหลาย นามสกุลไฟล์ (.jpg, .jpeg, .png, .JPG)
def get_tank_image_base64():
    possible_names = ["TANK.JPG", "TANK.jpg", "tank.jpg", "tank.jpeg", "TANK.PNG", "tank.png"]
    found_file = None
    
    for fname in possible_names:
        if os.path.exists(fname):
            found_file = fname
            break
            
    if found_file:
        with open(found_file, 'rb') as f:
            data = f.read()
        encoded = base64.b64encode(data).decode()
        ext = found_file.split('.')[-1].lower()
        mime_type = "image/png" if ext == "png" else "image/jpeg"
        return f"data:{mime_type};base64,{encoded}", found_file
    return None, None

bg_tank_url, filename_found = get_tank_image_base64()

# ตั้งค่า Fallback หากไม่พบรูปภาพ
bg_css = f"url('{bg_tank_url}')" if bg_tank_url else "linear-gradient(135deg, #1e293b 0%, #0f172a 100%)"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Kanit', sans-serif !important;
        background-color: #f8fafc !important;
    }}
    
    header[data-testid="stHeader"], footer {{ 
        display: none !important; 
    }}
    .main .block-container {{ 
        padding-top: 0rem !important; 
        padding-bottom: 2rem !important; 
        max-width: 95% !important; 
    }}

    /* Top Strip Header */
    .top-announcement {{
        background: #0f172a;
        color: #ffffff;
        text-align: center;
        padding: 6px;
        font-size: 0.85rem;
        letter-spacing: 1px;
        border-bottom: 3px solid #d32f2f;
        margin-bottom: 15px;
    }}

    /* Geometric Search Hero Card (หน้า HOME & SEARCH พร้อมภาพ TANK) */
    .tank-search-container {{
        background-image: {bg_css};
        background-size: cover;
        background-position: center;
        border-radius: 12px;
        padding: 40px;
        position: relative;
        overflow: hidden;
        box-shadow: 0 10px 25px rgba(0,0,0,0.15);
        border-left: 8px solid #d32f2f;
        margin-bottom: 25px;
    }}
    
    .tank-overlay {{
        background: rgba(15, 23, 42, 0.78);
        backdrop-filter: blur(6px);
        padding: 28px;
        border-radius: 8px;
        color: white;
    }}

    /* Geometric Accent Cards (กล่องทรงเรขาคณิต) */
    .geo-card {{
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-left: 5px solid #d32f2f;
        padding: 20px;
        border-radius: 6px;
        position: relative;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }}
    .geo-card::after {{
        content: "";
        position: absolute;
        top: 0;
        right: 0;
        width: 0;
        height: 0;
        border-style: solid;
        border-width: 0 20px 20px 0;
        border-color: transparent #d32f2f transparent transparent;
    }}
    
    .geo-metric-title {{
        font-size: 0.85rem;
        color: #64748b;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    .geo-metric-value {{
        font-size: 1.8rem;
        font-weight: 700;
        color: #0f172a;
        margin-top: 5px;
    }}

    /* Input Search Box Style */
    div[data-baseweb="input"] {{
        border: none !important;
        border-bottom: 2px solid #94a3b8 !important;
        border-radius: 0px !important;
        background: transparent !important;
    }}
    div[data-baseweb="input"]:focus-within {{
        border-bottom: 2px solid #d32f2f !important;
    }}
</style>
""", unsafe_allow_html=True)

# ==============================
# STEP 2: LOAD & CLEAN DATA
# ==============================
master_file = "Consum control.xlsx"
if os.path.exists(master_file):
    df_master = pd.read_excel(master_file)
else:
    st.error("❌ ไม่พบไฟล์ Consum control.xlsx ในโฟลเดอร์ระบบ")
    st.stop()

location_file_xlsx = "Location.xlsx"
location_file_csv = "Location.csv"
df_location = pd.read_excel(location_file_xlsx) if os.path.exists(location_file_xlsx) else (pd.read_csv(location_file_csv) if os.path.exists(location_file_csv) else None)

# Clean Columns
df_master.columns = df_master.columns.astype(str).str.replace('\n', ' ').str.strip()

moq_col = next((c for c in df_master.columns if "2569" in c), None) or "MOQ_Temp"
if moq_col == "MOQ_Temp": df_master["MOQ_Temp"] = 0

cost_col = next((c for c in df_master.columns if any(k in c.lower() for k in ["unit cost", "cost", "price", "ราคา"])), None) or "Cost_Temp"
if cost_col == "Cost_Temp": df_master["Cost_Temp"] = 0

cols_to_use = ["Item", "Item Description", "On hand", "Ss", "Ordered", "Min", "Max", moq_col, cost_col, "LT", "จับคู่"]
existing_master = [c for c in cols_to_use if c in df_master.columns]
master = df_master[existing_master].copy()

master.rename(columns={
    "Item": "Part Number", "Item Description": "Description", "On hand": "On Hand",
    "Ss": "SS", "Ordered": "Ordered", "Min": "Min", "Max": "Max",
    moq_col: "MOQ (ปี 2569)", cost_col: "Unit Cost (บาท)", "LT": "Lead Time", "จับคู่": "ABC-XYZ"
}, inplace=True)

master["Part Number"] = master["Part Number"].astype(str).str.strip()

for col in ["On Hand", "SS", "Ordered", "Min", "Max", "MOQ (ปี 2569)", "Unit Cost (บาท)", "Lead Time"]:
    if col in master.columns:
        master[col] = pd.to_numeric(master[col], errors='coerce').fillna(0)

# Merge Location
if df_location is not None:
    df_location.columns = df_location.columns.astype(str).str.replace('\n', ' ').str.strip()
    loc_item_col = df_location.columns[0]
    loc_val_col = df_location.columns[-1]
    df_location[loc_item_col] = df_location[loc_item_col].astype(str).str.strip()
    loc_summary = df_location[[loc_item_col, loc_val_col]].rename(columns={loc_item_col: "Part Number", loc_val_col: "Location"})
    master = master.merge(loc_summary, on="Part Number", how="left")
    master["Location"] = master["Location"].fillna("-")
else:
    master["Location"] = "-"

master["Total Stock"] = master["On Hand"] + master["Ordered"]

def get_status(r):
    if r["Total Stock"] < r["Min"]: return "🚨 Urgent"
    elif r["On Hand"] < r["Min"]: return "⚠️ Order"
    else: return "✅ Normal"

master["Status"] = master.apply(get_status, axis=1)
master["ยอดต้องสั่ง"] = master.apply(lambda r: int(max(r["Max"] - r["Total Stock"], r["MOQ (ปี 2569)"])) if r["Status"] in ["🚨 Urgent", "⚠️ Order"] and (r["Max"] - r["Total Stock"]) > 0 else 0, axis=1)
master["Total Cost (บาท)"] = master["ยอดต้องสั่ง"] * master["Unit Cost (บาท)"]

# ==============================
# STEP 3: HEADER (LOGO + SEARCH + ACTIONS)
# ==============================
st.markdown('<div class="top-announcement">HEIL TRAILER™ — EXECUTIVE INVENTORY & LOCATION CONTROL SYSTEM</div>', unsafe_allow_html=True)

col_logo, col_search, col_icons = st.columns([2.5, 5, 2.5])

with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=210)
    else:
        st.markdown("<h2 style='color:#d32f2f; margin:0; font-weight:800; font-style:italic;'>HEIL TRAILER™</h2>", unsafe_allow_html=True)

with col_search:
    search_input = st.text_input("Search", placeholder="🔍 ค้นหารายการสินค้า, Part Number หรือ Description...", label_visibility="collapsed")

with col_icons:
    st.markdown("""
        <div style="display: flex; justify-content: flex-end; gap: 20px; font-size: 0.8rem; color: #475569; text-align: center; margin-top: 5px;">
            <div>👤<br><b>ACCOUNT</b></div>
            <div>📍<br><b>LOCATION</b></div>
            <div>🚨<br><b>URGENT</b></div>
            <div>📋<br><b>MASTER</b></div>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

# Navigation Menu Bar
selected_menu = option_menu(
    menu_title=None,
    options=["HOME & SEARCH", "LOCATION FINDER", "URGENT ORDER", "MASTER DATA"],
    icons=["house", "geo-alt", "exclamation-circle", "collection"],
    default_index=0,
    orientation="horizontal",
    styles={
        "container": {"padding": "0!important", "background-color": "#ffffff", "border-top": "1px solid #e2e8f0", "border-bottom": "1px solid #e2e8f0", "border-radius": "0px"},
        "icon": {"color": "#64748b", "font-size": "13px"}, 
        "nav-link": {"font-size": "12px", "text-transform": "uppercase", "letter-spacing": "1px", "color": "#334155", "margin": "0px 15px", "padding": "12px 0px"},
        "nav-link-selected": {"background-color": "transparent", "color": "#d32f2f", "font-weight": "700", "border-bottom": "3px solid #d32f2f", "border-radius": "0px"},
    }
)

st.markdown("<br>", unsafe_allow_html=True)

# แจ้งเตือนสถานะการโหลดรูป TANK
if not bg_tank_url:
    st.warning("⚠️ ไม่พบไฟล์ภาพ TANK.JPG ในโฟลเดอร์ กรุณาตรวจสอบชื่อไฟล์ภาพในโฟลเดอร์ระบบ")

# ==============================
# STEP 4: PAGES & ROUTING
# ==============================

# --- MENU 1: HOME & SEARCH ---
if selected_menu == "HOME & SEARCH":
    if search_input:
        st.subheader(f"🔎 ผลการค้นหาสำหรับ: '{search_input}'")
        results = master[
            master["Part Number"].str.contains(search_input, case=False, na=False) |
            master["Description"].str.contains(search_input, case=False, na=False)
        ]
        st.dataframe(results[["Part Number", "Description", "On Hand", "Min", "Ordered", "Location", "Status"]], use_container_width=True, hide_index=True)
    else:
        # Hero Container พร้อมรูป TANK
        st.markdown("""
            <div class="tank-search-container">
                <div class="tank-overlay">
                    <h2 style="margin:0; font-weight:700; color:#ffffff; font-size:2rem;">HEIL TRAILER™ INVENTORY SEARCH</h2>
                    <p style="margin-top:8px; color:#cbd5e1; font-size:1rem;">พิมพ์ค้นหารายการอะไหล่ หรือ Part Number ที่ช่องด้านบนเพื่อตรวจสอบสต็อกและตำแหน่งจัดเก็บ</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        # กล่องทรงเรขาคณิต (Geometric KPI Cards)
        st.markdown("#### 📐 Geometric Inventory Summary")
        m1, m2, m3, m4 = st.columns(4)
        
        with m1:
            st.markdown(f"""
                <div class="geo-card">
                    <div class="geo-metric-title">TOTAL ITEMS</div>
                    <div class="geo-metric-value">{len(master):,} <span style="font-size:1rem; font-weight:normal;">Pcs</span></div>
                </div>
            """, unsafe_allow_html=True)
        with m2:
            st.markdown(f"""
                <div class="geo-card" style="border-left-color: #ef4444;">
                    <div class="geo-metric-title">URGENT ORDERS</div>
                    <div class="geo-metric-value" style="color: #ef4444;">{len(master[master['Status']=='🚨 Urgent']):,} <span style="font-size:1rem; font-weight:normal;">Pcs</span></div>
                </div>
            """, unsafe_allow_html=True)
        with m3:
            st.markdown(f"""
                <div class="geo-card" style="border-left-color: #f59e0b;">
                    <div class="geo-metric-title">REORDER POINT</div>
                    <div class="geo-metric-value" style="color: #f59e0b;">{len(master[master['Status']=='⚠️ Order']):,} <span style="font-size:1rem; font-weight:normal;">Pcs</span></div>
                </div>
            """, unsafe_allow_html=True)
        with m4:
            st.markdown(f"""
                <div class="geo-card" style="border-left-color: #10b981;">
                    <div class="geo-metric-title">LOCATED ITEMS</div>
                    <div class="geo-metric-value" style="color: #10b981;">{len(master[master['Location']!='-']):,} <span style="font-size:1rem; font-weight:normal;">Pcs</span></div>
                </div>
            """, unsafe_allow_html=True)

# --- MENU 2: LOCATION FINDER ---
elif selected_menu == "LOCATION FINDER":
    st.subheader("📍 ระบบค้นหาตำแหน่งจัดเก็บ (Location Finder)")
    loc_search = st.text_input("ค้นหาตำแหน่ง หรือ Part Number:", placeholder="พิมพ์ค้นหา Location หรือ Part Number...")
    
    df_loc_display = master.copy()
    if loc_search:
        df_loc_display = df_loc_display[
            df_loc_display["Part Number"].str.contains(loc_search, case=False, na=False) |
            df_loc_display["Location"].str.contains(loc_search, case=False, na=False)
        ]
    st.dataframe(df_loc_display[["Part Number", "Description", "Location", "On Hand", "Status"]], use_container_width=True, hide_index=True)

# --- MENU 3: URGENT ORDER ---
elif selected_menu == "URGENT ORDER":
    st.subheader("🚨 รายการสั่งซื้อด่วนและจุดสั่งเติม")
    to_order = master[master["ยอดต้องสั่ง"] > 0]
    st.dataframe(to_order[["Status", "Part Number", "Description", "Location", "ยอดต้องสั่ง", "MOQ (ปี 2569)", "Unit Cost (บาท)", "Total Cost (บาท)", "On Hand", "Min"]], use_container_width=True, hide_index=True)

# --- MENU 4: MASTER DATA ---
elif selected_menu == "MASTER DATA":
    st.subheader("📋 ข้อมูลคลังสินค้าทั้งหมด (Master Data)")
    st.dataframe(master, use_container_width=True, hide_index=True)