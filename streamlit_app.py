import streamlit as st
import struct
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import re

# ==========================================
# 1. Page Config & Custom CSS (Dark Mode)
# ==========================================
st.set_page_config(
    page_title="Recorder Furnace YOKOGAWA (.DAD)",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        /* ซ่อนแถบขาว Header ด้านบน */
        header[data-testid="stHeader"] {
            background-color: transparent !important;
            display: none !important;
        }
        [data-testid="stToolbar"] {
            display: none !important;
        }
        
        /* ตั้งค่าพื้นหลัง Dark Mode */
        html, body, .stApp, [data-testid="stAppViewContainer"] {
            background-color: #0e1117 !important;
            color: #ffffff !important;
        }
        [data-testid="stSidebar"] {
            background-color: #161b22 !important;
        }
        .stMarkdown, h1, h2, h3, p, span, label {
            color: #ffffff !important;
        }

        /* ปุ่มเคลียร์ข้อมูลใน Sidebar */
        [data-testid="stSidebar"] div.stButton > button {
            background-color: #21262d !important;
            color: #ffffff !important;
            border: 1px solid #F0B90B !important;
            font-weight: bold !important;
            width: 100% !important;
            padding: 8px 16px !important;
        }
        [data-testid="stSidebar"] div.stButton > button:hover {
            background-color: #F0B90B !important;
            color: #000000 !important;
        }

        /* กล่อง File Uploader */
        [data-testid="stFileUploader"] {
            background-color: #161b22 !important;
            border: 1.5px solid #F0B90B !important;
            border-radius: 8px !important;
            padding: 10px !important;
        }
        [data-testid="stFileUploader"] section {
            background-color: #1c2128 !important;
            border: 1px dashed #F0B90B !important;
            border-radius: 6px !important;
        }
        [data-testid="stFileUploader"] section div, 
        [data-testid="stFileUploader"] section span,
        [data-testid="stFileUploader"] section small {
            color: #e6edf3 !important;
        }

        /* การ์ดไฟล์ที่อัปโหลดแล้ว */
        [data-testid="stFileUploaderFileData"],
        [data-testid="stFileUploaderFileData"] > div,
        [data-testid="stFileUploaderFile"] {
            background-color: #21262d !important;
            border: 1px solid #F0B90B !important;
            border-radius: 6px !important;
        }
        [data-testid="stFileUploaderFileData"] *,
        [data-testid="stFileUploaderFile"] * {
            color: #ffffff !important;
            font-weight: bold !important;
        }
    </style>
""", unsafe_allow_html=True)

title_placeholder = st.empty()
title_placeholder.title("🏭 Recorder Furnace from YOKOGAWA (.DAD Data)")

# ==========================================
# 2. Robust DAD Parser Logic (แก้ปัญหาไฟล์ 028613)
# ==========================================
def extract_date_from_filename(filename):
    """ ถอดรหัสวันที่จากชื่อไฟล์เพื่อใช้เป็นตัวกรอง """
    match = re.search(r'(\d{2})(\d{2})(\d{2})_', filename)
    if match:
        yr, mo, dy = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 1 <= mo <= 12 and 1 <= dy <= 31:
            full_yr = 2000 + yr if yr < 80 else 1900 + yr
            return full_yr, mo, dy
    return None

def find_dad_params(raw):
    """ สแกนหา Offset อัจฉริยะ (เช็คซ้ำ 3 Records ป้องกันจับไบต์ผิดตำแหน่ง) """
    max_search = min(35000, len(raw) - 200)
    for offset in range(10000, max_search):
        yr, mo, dy, hr, mn, sc = raw[offset:offset+6]
        if 0 <= yr <= 99 and 1 <= mo <= 12 and 1 <= dy <= 31 and 0 <= hr <= 23 and 0 <= mn <= 59 and 0 <= sc <= 59:
            # ลองใช้ Record Size มาตรฐานต่างๆ ของ Yokogawa
            for rs in [84, 88, 92, 96, 80, 100, 128]:
                if offset + rs*2 + 6 < len(raw):
                    # เช็คข้อมูลล่วงหน้า 2-3 แถวให้แน่ใจว่าอ่านวันที่ได้ถูกต้อง (แก้บั๊กของไฟล์ 028613)
                    yr2, mo2, dy2, hr2, mn2, sc2 = raw[offset+rs : offset+rs+6]
                    yr3, mo3, dy3, hr3, mn3, sc3 = raw[offset+rs*2 : offset+rs*2+6]
                    if (0 <= yr2 <= 99 and 1 <= mo2 <= 12 and 1 <= dy2 <= 31 and 0 <= hr2 <= 23) and \
                       (0 <= yr3 <= 99 and 1 <= mo3 <= 12 and 1 <= dy3 <= 31 and 0 <= hr3 <= 23):
                        num_ch = (rs - 8) // 4
                        return offset, rs, num_ch
    return 15008, 88, 20

@st.cache_data(show_spinner="⏳ กำลังประมวลผลไฟล์ .DAD (Auto-Scale & MAX value)...")
def parse_dad_to_df(files_data):
    all_records = []
    machine_type = "NB1"
    
    for fname, raw in files_data:
        # เช็ครุ่นเตาจากชื่อไฟล์
        if "DATA" in fname.upper() or "_NB2" in fname.upper():
            machine_type = "NB2"
            
        file_date_hint = extract_date_from_filename(fname)
        offset, record_size, num_ch = find_dad_params(raw)
        
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
                
            parsed_yr = (2000 + yr) if yr < 80 else (1900 + yr)
            
            # บังคับปีให้ตรงกับชื่อไฟล์หาก header เสียหาย
            if file_date_hint:
                h_yr, h_mo, h_dy = file_date_hint
                if abs(parsed_yr - h_yr) > 5:
                    parsed_yr = h_yr
                    
            try:
                ts = datetime(parsed_yr, mo, dy, hr, mn, sc)
            except ValueError:
                continue
                
            rec = {'DateTime': ts}
            valid_data = False
            
            for ci in range(min(num_ch, 40)): 
                data_pos = base + 8 + ci*4
                if data_pos + 4 > len(raw): break
                
                # อ่านค่า MAX (2 ไบต์หลัง) จากแต่ละแชนแนล
                max_v = struct.unpack_from('>h', raw, data_pos + 2)[0]
                
                # กรองค่าขยะ และ error bytes
                if max_v not in (-32768, -32767, 32767, 32766) and (-30000 < max_v < 30000):
                    val = max_v / 10.0
                    if -100.0 <= val <= 2500.0:
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
        
        # จัด Map Channels
        for i in range(1, 8): df[f"Top Zone #{i}"] = df.get(f"CH{i:03d}")
        for i in range(1, 8): df[f"Bottom Zone #{i}"] = df.get(f"CH{i+7:03d}")
        
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
            
    return df, machine_type

# ==========================================
# 3. Chart Industrial Style (Auto Scaling)
# ==========================================
def apply_industrial_style(fig, y_title, is_dual_axis=False):
    layout_args = dict(
        template="plotly_dark",
        plot_bgcolor="#161b22",
        paper_bgcolor="#0e1117",
        hovermode="x unified",
        showlegend=True,
        legend=dict(
            font=dict(color="#FFFFFF", size=12, family="Arial Bold"),
            bgcolor="rgba(27, 31, 36, 0.95)",
            bordercolor="#F0B90B",
            borderwidth=1.5,
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        ),
        xaxis=dict(
            title=dict(text="Absolute Time [Date & Time]", font=dict(color="#FFFFFF", size=12)),
            tickfont=dict(color="#CCCCCC", size=10),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.08)",
            linecolor="#555555",
            type="date",
        ),
        yaxis=dict(
            title=dict(text=y_title, font=dict(color="#FFFFFF", size=12)),
            tickfont=dict(color="#CCCCCC", size=10),
            showgrid=True,
            gridcolor="rgba(255,255,255,0.08)",
            zeroline=False,
            linecolor="#555555",
            autorange=True  # เปลี่ยนเป็น Auto Range แทนการล็อคสเกล
        ),
        height=420,
        margin=dict(l=60, r=180, t=30, b=40),
    )
    fig.update_layout(**layout_args)

# ==========================================
# 4. Sidebar UI & Data Loading
# ==========================================
st.sidebar.header("📁 เมนูอัปโหลดข้อมูล (.DAD)")

if st.sidebar.button("🧹 เคลียร์ข้อมูลไฟล์เก่าทั้งหมด"):
    st.cache_data.clear()
    st.rerun()

uploaded_files = st.sidebar.file_uploader(
    "อัปโหลดไฟล์ YOKOGAWA (.DAD) ได้มากกว่า 1 ไฟล์", 
    type=["dad", "DAD"],
    accept_multiple_files=True
)

# ==========================================
# 5. Main Application Logic
# ==========================================
if uploaded_files:
    try:
        files_data = [(f.name, f.read()) for f in uploaded_files]
        df, detected_machine = parse_dad_to_df(files_data)
        
        if df.empty:
            st.warning("⚠️ ไม่พบข้อมูลที่สามารถอ่านได้ในไฟล์ที่อัปโหลด")
            st.stop()
            
        title_placeholder.title(f"🏭 Recorder {detected_machine} Furnace from YOKOGAWA (.DAD Data)")
        st.success(f"รวมข้อมูลสำเร็จ {len(uploaded_files)} ไฟล์ ({len(df)} แถว) - แสดงผลโดยใช้ค่า MAX (Auto-Scale)")

        # ---------------------------------------------------------
        # 1. Top Zone Temp (Auto Scale)
        # ---------------------------------------------------------
        st.subheader("1. Brazing zone Top #1-#7 (CH001-CH007)")
        fig1 = go.Figure()
        top_colors = ["#FF0000", "#008000", "#0000FF", "#8A2BE2", "#A52A2A", "#FFA500", "#9ACD32"]
        for i in range(1, 8):
            fig1.add_trace(go.Scatter(
                x=df["DateTime"], 
                y=df.get(f"Top Zone #{i}"), 
                name=f"Top Z#{i} (CH{i:03d})", 
                mode="lines", 
                line=dict(color=top_colors[i-1], width=2)
            ))
        apply_industrial_style(fig1, "Temperature (°C)")
        st.plotly_chart(fig1, use_container_width=True)

        # ---------------------------------------------------------
        # 2. Bottom Zone Temp (Auto Scale)
        # ---------------------------------------------------------
        st.subheader("2. Brazing zone Bottom #1-#7 (CH008-CH014)")
        fig2 = go.Figure()
        bottom_colors = ["#E0FFFF", "#FF1493", "#808080", "#00FF00", "#008000", "#0000FF", "#8A2BE2"]
        for i in range(1, 8):
            ch_num = 7 + i
            fig2.add_trace(go.Scatter(
                x=df["DateTime"], 
                y=df.get(f"Bottom Zone #{i}"), 
                name=f"Bottom Z#{i} (CH{ch_num:03d})", 
                mode="lines", 
                line=dict(color=bottom_colors[i-1], width=2)
            ))
        apply_industrial_style(fig2, "Temperature (°C)")
        st.plotly_chart(fig2, use_container_width=True)

        # ---------------------------------------------------------
        # 3. Dryer Temp (Auto Scale)
        # ---------------------------------------------------------
        st.subheader("3. Dryer #1 & #2 (CH016 & CH017)")
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(x=df["DateTime"], y=df.get("Dryer #1"), name="Dryer #1 (CH016)", mode="lines", line=dict(color="#FFA500", width=2)))
        fig3.add_trace(go.Scatter(x=df["DateTime"], y=df.get("Dryer #2"), name="Dryer #2 (CH017)", mode="lines", line=dict(color="#9ACD32", width=2)))
        apply_industrial_style(fig3, "Temperature (°C)")
        st.plotly_chart(fig3, use_container_width=True)

        # ---------------------------------------------------------
        # 4. O2 & N2 Flow Rate (Auto Scale)
        # ---------------------------------------------------------
        title_o2 = "4. ppmO2 Entry/Exit & N2 Flow (CH015, CH018, CH019)" if detected_machine == "NB1" else "4. ppmO2 Entry/Exit (CH015, CH018)"
        st.subheader(title_o2)
        fig4 = make_subplots(specs=[[{"secondary_y": True}]])
        
        ent_ch = "CH019" if detected_machine == "NB1" else "CH018"
        fig4.add_trace(go.Scatter(x=df["DateTime"], y=df.get("ENTRANCE O2"), name=f"ENTRANCE O2 ({ent_ch})", mode="lines", line=dict(color="#FF80FF", width=2)), secondary_y=False)
        fig4.add_trace(go.Scatter(x=df["DateTime"], y=df.get("EXIT O2"), name="EXIT O2 (CH015)", mode="lines", line=dict(color="#A52A2A", width=2)), secondary_y=False)
        
        if detected_machine == "NB1":
            fig4.add_trace(go.Scatter(x=df["DateTime"], y=df.get("N2 Flow"), name="N2 Flow (CH018)", mode="lines", line=dict(color="#ADD8E6", width=2)), secondary_y=True)
        
        apply_industrial_style(fig4, "Oxygen Level (ppm)", is_dual_axis=True)
        fig4.update_layout(
            yaxis=dict(
                title=dict(text="Oxygen Level (ppm)", font=dict(color="#FFFFFF", size=12)),
                autorange=True,  # ใช้ Auto Range
                showgrid=True,
                gridcolor="rgba(255,255,255,0.08)"
            ),
            yaxis2=dict(
                title=dict(text="N2 Flow Rate (m3/Hr)", font=dict(color="#ADD8E6", size=12)),
                tickfont=dict(color="#ADD8E6", size=10),
                showgrid=False,
                overlaying="y",
                side="right",
                linecolor="#ADD8E6",
                autorange=True,
                visible=(detected_machine == "NB1")
            )
        )
        st.plotly_chart(fig4, use_container_width=True)

        # ---------------------------------------------------------
        # 5. Dew Point (Auto Scale)
        # ---------------------------------------------------------
        if detected_machine == "NB1":
            st.subheader("5. Dew point 'Cdp (CH020)")
            fig5 = go.Figure()
            fig5.add_trace(go.Scatter(
                x=df["DateTime"], 
                y=df.get("DEW POINT"), 
                name="Dew Point (CH020)", 
                mode="lines", 
                line=dict(color="#00ecff", width=2)
            ))
            apply_industrial_style(fig5, "Dew Point (°Cdp)")
            st.plotly_chart(fig5, use_container_width=True)

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")

else:
    st.info("👈 กรุณาอัปโหลดไฟล์ .DAD ที่เมนูด้านซ้าย สามารถเลือกอัปโหลดพร้อมกันได้หลายไฟล์")
