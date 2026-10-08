import streamlit as st
import struct
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import io
import re

# ==========================================
# 1. Page Config & Default Setup
# ==========================================
st.set_page_config(
    page_title="Recorder Furnace YOKOGAWA (.DAD)",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "dad_uploader_key" not in st.session_state:
    st.session_state["dad_uploader_key"] = 0

# ==========================================
# 2. CSS Injector (ระบบจัดการ UI & Theme)
# ==========================================
st.markdown("""
    <style>
        [data-testid="stHeader"] { background-color: transparent !important; }
        [data-testid="stToolbar"], #MainMenu, .stAppDeployButton { display: none !important; } 
    </style>
""", unsafe_allow_html=True)

dark_css = """
<style>
    html, body, .stApp, [data-testid="stAppViewContainer"] { 
        background-color: #0e1117 !important; 
        color: #ffffff !important; 
    }
    [data-testid="stSidebar"], [data-testid="stSidebarHeader"] { background-color: #161b22 !important; }
    .stMarkdown, h1, h2, h3, h4, h5, h6, p, span, label { color: #ffffff !important; }
    
    button[kind="header"], [data-testid="collapsedControl"] { color: #F0B90B !important; }
    button[kind="header"] svg, [data-testid="collapsedControl"] svg { fill: #F0B90B !important; }

    div.stButton > button, [data-testid="stDownloadButton"] > button {
        background-color: #21262d !important;
        color: #ffffff !important;
        border: 1px solid #F0B90B !important;
        font-weight: bold !important;
        width: 100%;
    }
    div.stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        background-color: #F0B90B !important;
        color: #000000 !important;
    }
    div.stButton > button *, [data-testid="stDownloadButton"] > button * { color: inherit !important; }

    [data-testid="stFileUploader"] { background-color: #0e1117 !important; border: 1.5px solid #F0B90B !important; border-radius: 8px !important; padding: 10px !important; }
    [data-testid="stFileUploader"] section { background-color: #1c2128 !important; border: 1px dashed #F0B90B !important; }
    [data-testid="stFileUploaderFileData"] { background-color: #21262d !important; border: 1px solid #F0B90B !important; }
    
    [data-testid="stFileUploader"] button {
        background-color: #21262d !important;
        color: #ffffff !important;
        border: 1px solid #F0B90B !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #F0B90B !important;
        color: #000000 !important;
    }
</style>
"""

bright_css = """
<style>
    html, body, .stApp, [data-testid="stAppViewContainer"] { 
        background-color: #f4f6f9 !important; 
        color: #1a1a1a !important; 
    }
    [data-testid="stSidebar"], [data-testid="stSidebarHeader"] { background-color: #e9ecef !important; }
    .stMarkdown, h1, h2, h3, h4, h5, h6, p, span, label { color: #1a1a1a !important; }
    
    button[kind="header"], [data-testid="collapsedControl"] { color: #0056b3 !important; }
    button[kind="header"] svg, [data-testid="collapsedControl"] svg { fill: #0056b3 !important; }

    div.stButton > button, [data-testid="stDownloadButton"] > button {
        background-color: #ffffff !important;
        color: #0056b3 !important;
        border: 1.5px solid #0056b3 !important;
        font-weight: bold !important;
        width: 100%;
    }
    div.stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        background-color: #0056b3 !important;
        color: #ffffff !important;
    }
    div.stButton > button *, [data-testid="stDownloadButton"] > button * { color: inherit !important; }

    [data-testid="stFileUploader"] { background-color: #ffffff !important; border: 1.5px solid #0056b3 !important; border-radius: 8px !important; padding: 10px !important; }
    [data-testid="stFileUploader"] section { background-color: #f8f9fa !important; border: 1px dashed #0056b3 !important; }
    [data-testid="stFileUploaderFileData"] { background-color: #e9ecef !important; border: 1px solid #0056b3 !important; }
    
    [data-testid="stFileUploader"] button {
        background-color: #ffffff !important;
        color: #0056b3 !important;
        border: 1px solid #0056b3 !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #0056b3 !important;
        color: #ffffff !important;
    }
</style>
"""

system_css = """
<style>
    div.stButton > button, [data-testid="stDownloadButton"] > button { width: 100%; }
</style>
"""

# ==========================================
# 3. Sidebar UI (แผงควบคุมหลักด้านซ้าย)
# ==========================================
with st.sidebar:
    st.header("⚙️ แผงควบคุม (Controls)")
    theme_choice = st.radio("🎨 เลือกโทนสีหน้าจอ (Theme):", ["Dark", "Bright", "System"], index=0, horizontal=True)
    st.markdown("---")
    
    uploaded_files = st.file_uploader(
        "📁 อัปโหลดไฟล์ YOKOGAWA (.DAD)", 
        type=["dad", "DAD"],
        accept_multiple_files=True,
        key=f"dad_uploader_{st.session_state['dad_uploader_key']}" 
    )
    
    if st.button("🧹 เคลียร์ข้อมูลทั้งหมด"):
        st.cache_data.clear()
        st.session_state["dad_uploader_key"] += 1 
        st.rerun()

    export_placeholder = st.container()

if theme_choice == "Dark":
    st.markdown(dark_css, unsafe_allow_html=True)
elif theme_choice == "Bright":
    st.markdown(bright_css, unsafe_allow_html=True)
else:
    st.markdown(system_css, unsafe_allow_html=True)

# ==========================================
# 4. Main UI Header
# ==========================================
credit_color = "#8b949e" if theme_choice == "Dark" else "#6c757d" if theme_choice == "Bright" else "gray"

title_placeholder = st.empty()
title_placeholder.title("🏭 Recorder NB1 and NB2 Furnace from YOKOGAWA (.DAD Data)")

st.markdown(f"<p style='color: {credit_color}; font-size: 0.88rem; margin-top: -15px; margin-bottom: 15px;'><i>Wichien Laithanakit - Brazing Engineer - VSTS / Power Chonburi</i></p>", unsafe_allow_html=True)

file_names_placeholder = st.empty() 
st.markdown("---")

# ==========================================
# 5. DAD Parser Logic (อ่านค่า MAX + ตรวจสอบ NB1/NB2)
# ==========================================
def extract_date_from_filename(filename):
    match = re.search(r'_(\d{2})(\d{2})(\d{2})_', filename)
    if match:
        yr, mo, dy = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 1 <= mo <= 12 and 1 <= dy <= 31:
            full_yr = 2000 + yr if yr < 80 else 1900 + yr
            return full_yr, mo, dy
    return None

def find_dad_params(raw, machine_type):
    default_rs = 84 if machine_type == "NB2" else 88
    default_offset = 15008
    if default_offset + default_rs * 2 < len(raw):
        yr, mo, dy, hr, mn, sc = raw[default_offset:default_offset+6]
        if 0 <= yr <= 99 and 1 <= mo <= 12 and 1 <= dy <= 31 and 0 <= hr <= 23 and 0 <= mn <= 59 and 0 <= sc <= 59:
            return default_offset, default_rs, (default_rs - 8) // 4
            
    max_search = min(30000, len(raw) - 200)
    for offset in range(14000, max_search):
        yr, mo, dy, hr, mn, sc = raw[offset], raw[offset+1], raw[offset+2], raw[offset+3], raw[offset+4], raw[offset+5]
        if 0 <= yr <= 99 and 1 <= mo <= 12 and 1 <= dy <= 31 and 0 <= hr <= 23 and 0 <= mn <= 59 and 0 <= sc <= 59:
            for rs in [84, 88, 92, 80, 96]:
                if offset + rs + 5 < len(raw):
                    yr2, mo2, dy2 = raw[offset+rs], raw[offset+rs+1], raw[offset+rs+2]
                    if 0 <= yr2 <= 99 and 1 <= mo2 <= 12 and 1 <= dy2 <= 31:
                        return offset, rs, (rs - 8) // 4
    return default_offset, default_rs, (default_rs - 8) // 4

@st.cache_data(show_spinner="⏳ กำลังประมวลผลไฟล์ .DAD (ดึงเฉพาะค่า MAX)...")
def parse_dad_to_df(files_data):
    all_records = []
    machine_type = "Unknown"
    
    for fname, raw in files_data:
        # แยกเตาอัตโนมัติจากชื่อไฟล์
        if "_DATA" in fname.upper():
            machine_type = "NB2"
        else:
            machine_type = "NB1"
            
        file_date_hint = extract_date_from_filename(fname)
        offset, record_size, num_ch = find_dad_params(raw, machine_type)
        body_len = len(raw) - offset
        if body_len <= 0: continue
        total = body_len // record_size
        
        for i in range(total):
            base = offset + i * record_size
            hdr = raw[base:base+8]
            if len(hdr) < 8: break
            yr, mo, dy, hr, mn, sc = hdr[0], hdr[1], hdr[2], hdr[3], hdr[4], hdr[5]
            
            if not (0 <= yr <= 99 and 1 <= mo <= 12 and 1 <= dy <= 31 and 0 <= hr <= 23 and 0 <= mn <= 59 and 0 <= sc <= 59):
                continue
                
            full_year = (2000 + yr) if yr < 80 else (1900 + yr)
            if file_date_hint:
                hint_yr, hint_mo, hint_dy = file_date_hint
                if abs(full_year - hint_yr) > 2:
                    full_year, mo, dy = hint_yr, hint_mo, hint_dy
                    
            try:
                ts = datetime(full_year, mo, dy, hr, mn, sc)
            except ValueError:
                continue
                
            rec = {'DateTime': ts}
            valid_data = False
            for ci in range(min(num_ch, 40)): 
                data_pos = base + 8 + ci*4
                if data_pos + 4 > len(raw): break
                
                # อ่านค่า MAX (2 ไบต์หลังของ 4 ไบต์)
                max_v = struct.unpack_from('>h', raw, data_pos + 2)[0]
                
                # กรองค่า Error / Sensor Out of range
                if max_v not in (-32768, -32767, 32767) and (-30000 < max_v < 30000):
                    val = max_v / 10.0
                    if -100.0 <= val <= 2000.0:
                        rec[f'CH{(ci+1):03d}'] = val
                        valid_data = True
                    else:
                        rec[f'CH{(ci+1):03d}'] = None
                else:
                    rec[f'CH{(ci+1):03d}'] = None
                    
            if valid_data:
                all_records.append(rec)
                
    df = pd.DataFrame(all_records)
    if not df.empty:
        df = df.drop_duplicates(subset=["DateTime"]).sort_values("DateTime").reset_index(drop=True)
        
        # กรองเฉพาะช่วงเวลาหลัก
        median_dt = df["DateTime"].iloc[len(df)//2]
        df = df[abs(df["DateTime"] - median_dt) <= pd.Timedelta(days=10)]
        
        # Map Channel ตรงตามสเปกของรูปภาพ
        for i in range(1, 8): df[f"Top Zone #{i}"] = df.get(f"CH{i:03d}")        # CH001-CH007
        for i in range(1, 8): df[f"Bottom Zone #{i}"] = df.get(f"CH{(i+7):03d}")  # CH008-CH014
        
        df["EXIT O2"] = df.get("CH015")
        df["Dryer #1"] = df.get("CH016")
        df["Dryer #2"] = df.get("CH017")
        
        if machine_type == "NB1":
            df["ENTRANCE O2"] = df.get("CH019")
            df["N2 Flow"] = df.get("CH018")
            df["DEW POINT"] = df.get("CH020")
        else:
            df["ENTRANCE O2"] = df.get("CH018")
            df["N2 Flow"] = None
            df["DEW POINT"] = None
            
        df = df.sort_values("DateTime").reset_index(drop=True)
        
    return df, machine_type

# ==========================================
# 6. Industrial Style & Chart Creator
# ==========================================
def apply_industrial_style(fig, y_title, theme_mode, is_dual_axis=False):
    t_bg = "#161b22" if theme_mode == "Dark" else "#ffffff" if theme_mode == "Bright" else "rgba(0,0,0,0)"
    t_paper = "#0e1117" if theme_mode == "Dark" else "#f4f6f9" if theme_mode == "Bright" else "rgba(0,0,0,0)"
    t_font = "#ffffff" if theme_mode == "Dark" else "#1a1a1a" if theme_mode == "Bright" else "gray"
    t_grid = "rgba(255,255,255,0.08)" if theme_mode == "Dark" else "rgba(0,0,0,0.1)"
    t_line = "#555555" if theme_mode == "Dark" else "#cccccc"

    layout_args = dict(
        plot_bgcolor=t_bg,
        paper_bgcolor=t_paper,
        hovermode="x unified",
        showlegend=True,
        legend=dict(
            font=dict(color=t_font, size=12, family="Arial Bold"),
            bgcolor="rgba(128, 128, 128, 0.1)",
            bordercolor=t_line,
            borderwidth=1.5,
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        ),
        xaxis=dict(
            title=dict(text="Absolute Time [Date & Time]", font=dict(color=t_font, size=12)),
            tickfont=dict(color=t_font, size=10),
            showgrid=True,
            gridcolor=t_grid,
            linecolor=t_line,
            type="date",
        ),
        yaxis=dict(
            title=dict(text=y_title, font=dict(color=t_font, size=12)),
            tickfont=dict(color=t_font, size=10),
            showgrid=True,
            gridcolor=t_grid,
            zeroline=False,
            linecolor=t_line,
        ),
        height=480, 
    )
    fig.update_layout(**layout_args)
    return t_font

def create_unified_figure(df, machine_type, initial_idx, theme_mode):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    traces_info = []
    
    top_colors = ["#FF0000", "#008000", "#0000FF", "#8A2BE2", "#A52A2A", "#FFA500", "#9ACD32"]
    bottom_colors = ["#00FFFF", "#FF1493", "#808080", "#00FF00", "#008000", "#0000FF", "#8A2BE2"]

    # Top Zone (CH001-CH007)
    for i in range(1, 8):
        fig.add_trace(go.Scatter(x=df["DateTime"], y=df.get(f"Top Zone #{i}"), name=f"Top Z#{i} (CH{i:03d})", mode="lines", line=dict(color=top_colors[i-1], width=2), visible=(initial_idx==0)), secondary_y=False)
        traces_info.append(0)

    # Bottom Zone (CH008-CH014)
    for i in range(1, 8):
        ch_num = 7 + i
        fig.add_trace(go.Scatter(x=df["DateTime"], y=df.get(f"Bottom Zone #{i}"), name=f"Bottom Z#{i} (CH{ch_num:03d})", mode="lines", line=dict(color=bottom_colors[i-1], width=2), visible=(initial_idx==1)), secondary_y=False)
        traces_info.append(1)

    # Dryer
    fig.add_trace(go.Scatter(x=df["DateTime"], y=df.get("Dryer #1"), name="Dryer #1 (CH016)", mode="lines", line=dict(color="#FFA500", width=2), visible=(initial_idx==2)), secondary_y=False)
    traces_info.append(2)
    fig.add_trace(go.Scatter(x=df["DateTime"], y=df.get("Dryer #2"), name="Dryer #2 (CH017)", mode="lines", line=dict(color="#9ACD32", width=2), visible=(initial_idx==2)), secondary_y=False)
    traces_info.append(2)

    # O2 / N2
    ent_ch = "CH019" if machine_type == "NB1" else "CH018"
    fig.add_trace(go.Scatter(x=df["DateTime"], y=df.get("ENTRANCE O2"), name=f"ENTRANCE O2 ({ent_ch})", mode="lines", line=dict(color="#FF80FF", width=2), visible=(initial_idx==3)), secondary_y=False)
    traces_info.append(3)
    fig.add_trace(go.Scatter(x=df["DateTime"], y=df.get("EXIT O2"), name="EXIT O2 (CH015)", mode="lines", line=dict(color="#A52A2A", width=2), visible=(initial_idx==3)), secondary_y=False)
    traces_info.append(3)

    if machine_type == "NB1" and "N2 Flow" in df.columns:
        fig.add_trace(go.Scatter(x=df["DateTime"], y=df.get("N2 Flow"), name="N2 Flow (CH018)", mode="lines", line=dict(color="#ADD8E6", width=2), visible=(initial_idx==3)), secondary_y=True)
        traces_info.append(3)

    # Dew Point
    if machine_type == "NB1" and "DEW POINT" in df.columns:
        fig.add_trace(go.Scatter(x=df["DateTime"], y=df.get("DEW POINT"), name="Dew Point (CH020)", mode="lines", line=dict(color="#00ecff", width=2), visible=(initial_idx==4)), secondary_y=False)
        traces_info.append(4)

    buttons = []
    buttons.append(dict(label="1. Top Zone", method="update", args=[{"visible": [t == 0 for t in traces_info]}, {"title.text": "<b>1. Brazing zone Top #1-#7 (CH001-CH007)</b>", "yaxis.title.text": "Temperature (°C)", "yaxis.range": [550, 650], "yaxis2.visible": False}]))
    buttons.append(dict(label="2. Bottom Zone", method="update", args=[{"visible": [t == 1 for t in traces_info]}, {"title.text": "<b>2. Brazing zone Bottom #1-#7 (CH008-CH014)</b>", "yaxis.title.text": "Temperature (°C)", "yaxis.range": [550, 650], "yaxis2.visible": False}]))
    buttons.append(dict(label="3. Dryer", method="update", args=[{"visible": [t == 2 for t in traces_info]}, {"title.text": "<b>3. Dryer #1 & #2 (CH016 & CH017)</b>", "yaxis.title.text": "Temperature (°C)", "yaxis.range": [150, 350], "yaxis2.visible": False}]))
    
    title_o2 = "<b>4. ppmO2 Entry/Exit & N2 Flow (CH015, CH018, CH019)</b>" if machine_type == "NB1" else "<b>4. ppmO2 Entry/Exit (CH015, CH018)</b>"
    buttons.append(dict(label="4. O2/N2", method="update", args=[{"visible": [t == 3 for t in traces_info]}, {"title.text": title_o2, "yaxis.title.text": "Oxygen Level (ppm) [0-200]", "yaxis.range": [0, 200], "yaxis2.visible": (machine_type == "NB1")}]))

    if machine_type == "NB1":
        buttons.append(dict(label="5. Dew Point", method="update", args=[{"visible": [t == 4 for t in traces_info]}, {"title.text": "<b>5. Dew point 'Cdp (CH020)</b>", "yaxis.title.text": "Dew Point (°Cdp)", "yaxis.range": [-100, 10], "yaxis2.visible": False}]))

    btn_bg = "rgba(22, 27, 34, 0.8)" if theme_mode == "Dark" else "rgba(255, 255, 255, 0.9)" if theme_mode == "Bright" else "rgba(128, 128, 128, 0.2)"
    btn_font = "#FFFFFF" if theme_mode == "Dark" else "#0056b3" if theme_mode == "Bright" else "gray"
    btn_border = "rgba(240, 185, 11, 0.5)" if theme_mode == "Dark" else "rgba(0, 86, 179, 0.5)" if theme_mode == "Bright" else "rgba(128, 128, 128, 0.5)"

    title_font_color = apply_industrial_style(fig, "Temperature (°C)", theme_mode, is_dual_axis=True)
    
    initial_title = buttons[initial_idx]["args"][1]["title.text"]
    initial_y_title = buttons[initial_idx]["args"][1]["yaxis.title.text"]
    initial_y_range = buttons[initial_idx]["args"][1].get("yaxis.range", None)
    initial_y2_visible = buttons[initial_idx]["args"][1]["yaxis2.visible"]

    fig.update_layout(
        title=dict(
            text=initial_title, 
            font=dict(size=18, color=title_font_color),
            x=0.0,
            y=0.98,
            xref="paper",
            yref="container", 
            xanchor="left",
            yanchor="top"
        ),
        margin=dict(l=60, r=100, t=130, b=40), 
        yaxis=dict(title=dict(text=initial_y_title), range=initial_y_range),
        yaxis2=dict(
            title=dict(text="N2 Flow Rate (m3/Hr)", font=dict(color="#ADD8E6", size=12)),
            tickfont=dict(color="#ADD8E6", size=10),
            showgrid=False,
            overlaying="y",
            side="right",
            linecolor="#ADD8E6",
            visible=initial_y2_visible,
            autorange=True
        ),
        updatemenus=[
            dict(
                type="buttons",
                direction="right",
                showactive=False,  
                x=0.0,
                y=1.15,
                xanchor="left",
                yanchor="bottom",
                buttons=buttons,
                font=dict(color=btn_font, size=11, family="Arial Bold"),
                bgcolor=btn_bg,  
                bordercolor=btn_border,
            )
        ]
    )
    return fig

# ==========================================
# 7. Main Application Logic
# ==========================================
if uploaded_files:
    try:
        detected_types = set()
        for f in uploaded_files:
            if "_DATA" in f.name.upper():
                detected_types.add("NB2")
            else:
                detected_types.add("NB1")
                
        if len(detected_types) > 1:
            st.error("❌ **ข้อผิดพลาด:** ตรวจพบการอัปโหลดไฟล์จากเตาคนละรุ่น (NB1 ผสมกับ NB2) กรุณากดปุ่ม 'เคลียร์ข้อมูลทั้งหมด' แล้วเลือกไฟล์จากเตาเดียวกันเท่านั้น")
            st.stop()

        files_data = [(f.name, f.read()) for f in uploaded_files]
        df, detected_machine = parse_dad_to_df(files_data)
        
        if df.empty:
            st.warning("⚠️ ไม่พบข้อมูลที่สามารถอ่านได้ในไฟล์ที่อัปโหลด")
            st.stop()
            
        title_placeholder.title(f"🏭 Recorder {detected_machine} Furnace from YOKOGAWA (.DAD Data)")
        file_names_str = ", ".join([f.name for f in uploaded_files])
        
        subtext_color = "#a0aab2" if theme_choice == "Dark" else "#666666" if theme_choice == "Bright" else "gray"
        file_names_placeholder.markdown(f"<span style='color:{subtext_color}; font-size:1.1rem;'><b>📁 File(s):</b> {file_names_str}</span>", unsafe_allow_html=True)
        
        st.success(f"รวมข้อมูลสำเร็จ {len(uploaded_files)} ไฟล์ ({len(df)} แถว) - โหมด {detected_machine} (ดึงเฉพาะค่า MAX)")
        
        # แสดงผลกราฟ
        st.plotly_chart(create_unified_figure(df, detected_machine, 0, theme_choice), use_container_width=True)
        st.plotly_chart(create_unified_figure(df, detected_machine, 1, theme_choice), use_container_width=True)
        st.plotly_chart(create_unified_figure(df, detected_machine, 2, theme_choice), use_container_width=True)
        st.plotly_chart(create_unified_figure(df, detected_machine, 3, theme_choice), use_container_width=True)
        
        if detected_machine == "NB1":
            st.plotly_chart(create_unified_figure(df, detected_machine, 4, theme_choice), use_container_width=True)

        # ==========================================
        # 8. ส่วนดาวน์โหลดข้อมูล (แสดงใน Sidebar)
        # ==========================================
        with export_placeholder:
            st.markdown("---")
            st.subheader("💾 ส่งออกข้อมูล (Export Data)")
            dl_type = st.radio("เลือกชนิดไฟล์ (Select Format):", ["CSV (.csv)", "Excel (.xlsx)"])
            st.write("") 
            
            if dl_type == "CSV (.csv)":
                csv_data = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 ดาวน์โหลดข้อมูล (Download)",
                    data=csv_data,
                    file_name=f"Furnace_{detected_machine}_Data.csv",
                    mime="text/csv"
                )
            else:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Data', index=False)
                st.download_button(
                    label="📥 ดาวน์โหลดข้อมูล (Download)",
                    data=buffer.getvalue(),
                    file_name=f"Furnace_{detected_machine}_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")

else:
    st.info("👈 กรุณาเปิดแถบควบคุมด้านซ้ายมือ (Sidebar) เพื่ออัปโหลดไฟล์ .DAD (สามารถอัปโหลดพร้อมกันได้หลายไฟล์)")
