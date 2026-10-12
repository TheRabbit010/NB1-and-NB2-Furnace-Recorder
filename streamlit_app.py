import streamlit as st
import struct
import io
import re
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title='DAD Report — Valeo NB1 & NB2',
    page_icon='🏭',
    layout='wide',
    initial_sidebar_state='expanded',
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F4E79, #2E75B6);
        color: white; padding: 20px 28px; border-radius: 12px;
        margin-bottom: 24px;
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; }
    .main-header p  { margin: 4px 0 0; opacity: 0.85; font-size: 0.9rem; }
    .stat-box {
        border: 1px solid var(--secondary-background-color);
        background-color: var(--secondary-background-color);
        border-radius: 10px; padding: 14px 18px; text-align: center;
    }
    .stat-box .val { font-size: 1.5rem; font-weight: 700; color: #2E75B6; }
    .stat-box .lbl { font-size: 0.78rem; opacity: 0.8; margin-top: 2px; }
    
    /* แต่งให้กล่อง Multiselect ดูโปร่งและสะอาดขึ้น (ทับสีแดง) */
    div[data-baseweb="select"] > div {
        background-color: rgba(128, 128, 128, 0.1); 
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
ALL_CH = [
    ('CH001','Top1 (Z1)',      '°C'),
    ('CH002','Top2 (Z2)',      '°C'),
    ('CH003','Top3 (Z3)',      '°C'),
    ('CH004','Top4 (Z4)',      '°C'),
    ('CH005','Top5 (Z5)',      '°C'),
    ('CH006','Top6 (Z6)',      '°C'),
    ('CH007','Top7 (Z7)',      '°C'),
    ('CH008','Bot1 (Z8)',      '°C'),
    ('CH009','Bot2 (Z9)',      '°C'),
    ('CH010','Bot3 (Z10)',     '°C'),
    ('CH011','Bot4 (Z11)',     '°C'),
    ('CH012','Bot5 (Z12)',     '°C'),
    ('CH013','Bot6 (Z13)',     '°C'),
    ('CH014','Bot7 (Z14)',     '°C'),
    ('CH015','O2 Exit',        'ppm'),
    ('CH016','Dryer zone1',    '°C'),
    ('CH017','Dryer zone2',    '°C'),
    ('CH018','O2 Entrance',    'ppm'),
]

COLORS_TOP = ['#FF0000', '#008000', '#0000FF', '#8A2BE2', '#A52A2A', '#FFA500', '#9ACD32'] 
COLORS_BOT = ['#ADD8E6', '#FF69B4', '#808080', '#FF0000', '#008000', '#0000FF', '#8A2BE2']
COLORS_DRY = {'CH016':'#FFA500', 'CH017':'#9ACD32'} 
COLORS_O2  = {'CH015':'#A52A2A', 'CH018':'#ADD8E6'}

# ── Robust Parser ─────────────────────────────────────────────────────────────
def find_dad_params(raw):
    max_search = min(35000, len(raw) - 200)
    for offset in range(10000, max_search):
        yr, mo, dy, hr, mn, sc = raw[offset:offset+6]
        if 0 <= yr <= 99 and 1 <= mo <= 12 and 1 <= dy <= 31 and 0 <= hr <= 23 and 0 <= mn <= 59 and 0 <= sc <= 59:
            for rs in [84, 88, 92, 96, 80, 100, 128]:
                if offset + rs*2 + 6 < len(raw):
                    yr2, mo2, dy2, hr2, mn2, sc2 = raw[offset+rs : offset+rs+6]
                    yr3, mo3, dy3, hr3, mn3, sc3 = raw[offset+rs*2 : offset+rs*2+6]
                    if (0 <= yr2 <= 99 and 1 <= mo2 <= 12 and 1 <= dy2 <= 31 and 0 <= hr2 <= 23) and \
                       (0 <= yr3 <= 99 and 1 <= mo3 <= 12 and 1 <= dy3 <= 31 and 0 <= hr3 <= 23):
                        return offset, rs
    return 15008, 84

def parse_dad(raw: bytes) -> list:
    records = []
    offset, record_size = find_dad_params(raw)
    
    body_len = len(raw) - offset
    if body_len <= 0: return records
    total = body_len // record_size
    
    for i in range(total):
        base = offset + i * record_size
        hdr  = raw[base:base+8]
        if len(hdr) < 8: break
        yr, mo, dy, hr, mn, sc = hdr[0], hdr[1], hdr[2], hdr[3], hdr[4], hdr[5]
        
        if not (0 <= yr <= 99 and 1 <= mo <= 12 and 1 <= dy <= 31 and 0 <= hr <= 23 and 0 <= mn <= 59 and 0 <= sc <= 59):
            continue
            
        parsed_yr = (2000 + yr) if yr < 80 else (1900 + yr)
        try:
            ts = datetime(parsed_yr, mo, dy, hr, mn, sc)
        except ValueError:
            continue
            
        rec = {'ts': ts}
        for ci, (ch, _, _) in enumerate(ALL_CH):
            data_pos = base + 8 + ci*4
            if data_pos + 4 > len(raw): break
            
            min_v = struct.unpack_from('>h', raw, data_pos)[0]
            max_v = struct.unpack_from('>h', raw, data_pos + 2)[0]
            
            def is_valid(v):
                return v not in (-32768, -32767, 32767, 32766) and (-30000 < v < 30000)
            
            if is_valid(min_v) and is_valid(max_v):
                min_val = min_v / 10.0
                max_val = max_v / 10.0
                if -100.0 <= min_val <= 2000.0 and -100.0 <= max_val <= 2000.0:
                    rec[f'{ch}_min'] = round(min_val, 1)
                    rec[f'{ch}_max'] = round(max_val, 1)
                    rec[f'{ch}_avg'] = round((min_val + max_val) / 2.0, 1)
                    rec[ch] = rec[f'{ch}_min']
                else:
                    rec[f'{ch}_min'] = rec[f'{ch}_max'] = rec[f'{ch}_avg'] = rec[ch] = None
            else:
                rec[f'{ch}_min'] = rec[f'{ch}_max'] = rec[f'{ch}_avg'] = rec[ch] = None
                
        records.append(rec)
    return records

# ── Excel Builder ─────────────────────────────────────────────────────────────
def build_excel(data, dt_start, dt_end, zones_sel):
    hdr_fill  = PatternFill('solid', fgColor='1F4E79')
    hdr_font  = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    norm_font = Font(name='Arial', size=9)
    thin      = Side(border_style='thin', color='C0C8D0')
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hcell(ws, col, row, val):
        c = ws.cell(row, col, val)
        c.fill=hdr_fill; c.font=hdr_font; c.border=border
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        return c

    def dcell(ws, col, row, val, fill=None, fmt=None, bold=False):
        c = ws.cell(row, col, val)
        c.fill=fill or PatternFill('solid',fgColor='FFFFFF')
        c.font=Font(name='Arial',size=9,bold=bold); c.border=border
        c.alignment=Alignment(horizontal='right' if isinstance(val,(int,float)) else 'left')
        if fmt: c.number_format=fmt
        return c

    sel_ch = [(ch,name,unit) for ch,name,unit in ALL_CH if name in zones_sel]
    wb = Workbook()
    n  = len(data)
    ds, de = 3, n+2
    report_name = f'{dt_start.strftime("%d-%m-%Y %H.%M")} to {dt_end.strftime("%d-%m-%Y %H.%M")}'

    ws = wb.active; ws.title='Raw Data'
    ws['A1'] = f'DAD Report — {report_name} | {n:,} records'
    ws.merge_cells(f'A1:{get_column_letter(3+len(sel_ch))}1')
    ws['A1'].font=Font(bold=True,name='Arial',size=12,color='1F4E79')
    ws.row_dimensions[1].height=22
    hcell(ws,1,2,'Timestamp'); hcell(ws,2,2,'Date'); hcell(ws,3,2,'Time')
    for i,(ch,name,unit) in enumerate(sel_ch):
        hcell(ws,4+i,2,f'{name}\n({ch})\n[{unit}]')
    ws.row_dimensions[2].height=48
    
    for ri,rec in enumerate(data,start=3):
        ts=rec['ts']; dp=ts.strftime('%Y/%m/%d'); tp=ts.strftime('%H:%M:%S')
        bg=PatternFill('solid',fgColor='FFFFFF')
        dcell(ws,1,ri,ts.strftime('%Y/%m/%d %H:%M:%S'),bg)
        dcell(ws,2,ri,dp,bg); dcell(ws,3,ri,tp,bg)
        for i,(ch,_,_) in enumerate(sel_ch):
            c=ws.cell(ri,4+i,rec[ch]); c.fill=bg; c.font=norm_font
            c.border=border; c.alignment=Alignment(horizontal='right'); c.number_format='0.0'
            
    ws.column_dimensions['A'].width=22; ws.column_dimensions['B'].width=13; ws.column_dimensions['C'].width=10
    for i in range(len(sel_ch)): ws.column_dimensions[get_column_letter(4+i)].width=14
    ws.freeze_panes='A3'; ws.auto_filter.ref=f'A2:{get_column_letter(3+len(sel_ch))}{n+2}'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── Chart Builder ─────────────────────────────────────────────────────────────
# 1. ฟังก์ชันช่วยสำหรับสร้างกราฟ 1 ตัว แยกออกจากกัน
def render_single_chart(data, sel_names, all_channels_ref, colors_ref, title, y_range, y_title):
    if not sel_names:
        return None
        
    sel_ch_info = [(ch,n,u) for ch,n,u in all_channels_ref if n in sel_names]
    if not sel_ch_info:
        return None

    fig = go.Figure()
    ts = [r['ts'] for r in data]
    
    for i, (ch, name, unit) in enumerate(sel_ch_info):
        color = colors_ref[i] if isinstance(colors_ref, list) else colors_ref.get(ch, '#888')
        fig.add_trace(go.Scatter(
            x=ts, y=[r[ch] for r in data],
            name=name, mode='lines',
            line=dict(color=color, width=1.5),
            hovertemplate=f'<b>{name}</b>: %{{y:.1f}} {unit}<extra></extra>',
        ))

    fig.update_layout(
        title=title,
        height=320,
        hovermode='x unified',
        hoverlabel=dict(font_size=11, namelength=30),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
            borderwidth=1, font=dict(size=11)
        ),
        margin=dict(l=50, r=20, t=50, b=20),
    )
    
    fig.update_yaxes(range=y_range, title_text=y_title)
    fig.update_xaxes(
        tickformat='%d/%m\n%H:%M',
        dtick=1*3600*1000, 
        tickfont=dict(size=10),
        showticklabels=True,
    )
    return fig

# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════
ICON_B64 = "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"

# 2. แก้ไข Header เพิ่ม NB2
st.markdown(f"""
<div class="main-header">
    <h1><img src="data:image/png;base64,{ICON_B64}" style="height:48px;vertical-align:middle;margin-right:10px;border-radius:10px"> DAD Report Generator</h1>
    <p>Valeo — NB1 & NB2 Furnace Data Analysis</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown('### 📂 Upload ไฟล์ .DAD')
    uploaded = st.file_uploader(
        'ลากไฟล์ .DAD มาวางที่นี่ (เลือกได้หลายไฟล์)',
        type=['DAD','dad'],
        accept_multiple_files=True,
    )
    
    st.markdown("---")
    st.markdown('### ⚙️ ตั้งค่าการแสดงผลกราฟ')
    
    # 3. เปลี่ยนจาก Radio เป็น Dropdown (Selectbox) และตั้งค่า MAX เป็นค่าเริ่มต้น
    value_mode = st.selectbox(
        'เลือกค่าที่ต้องการแสดงผล:',
        options=['MIN', 'AVG (เฉลี่ย)', 'MAX'],
        index=2 # 0:MIN, 1:AVG, 2:MAX
    )
    mode_key = {'MIN': '_min', 'AVG (เฉลี่ย)': '_avg', 'MAX': '_max'}[value_mode]

if not uploaded:
    st.info('👉 กรุณา Upload ไฟล์ .DAD จากแถบเมนูด้านซ้ายเพื่อเริ่มสร้างรายงานครับ')
    st.stop()

# 1. ย้ายกล่องเลือกตัวแปรไปอยู่ในฟังก์ชันสร้างกราฟเพื่อความสวยงาม
all_names_top = [name for _,name,_ in ALL_CH[:7]]
all_names_bot = [name for _,name,_ in ALL_CH[7:14]]
all_names_dry = [name for _,name,_ in ALL_CH[15:17]]
all_names_o2  = [name for _,name,_ in [ALL_CH[14], ALL_CH[17]]]

@st.cache_data(show_spinner='⏳ กำลังประมวลผลไฟล์...')
def load_data_auto(files_data, mode_k):
    all_records = []
    for fname, raw in files_data:
        recs = parse_dad(raw)
        all_records.extend(recs)
        
    if not all_records: return []
    
    all_records.sort(key=lambda x: x['ts'])
    seen = set()
    unique = []
    for r in all_records:
        k = r['ts']
        if k not in seen:
            seen.add(k)
            new_r = dict(r)
            for ch,_,_ in ALL_CH:
                new_r[ch] = r.get(f'{ch}{mode_k}')
            unique.append(new_r)
    return unique

files_data = [(f.name, f.read()) for f in uploaded]
data = load_data_auto(tuple((n,d) for n,d in files_data), mode_key)

if len(data) == 0:
    st.error('🚨 ไม่พบข้อมูลที่อ่านได้จากไฟล์ที่อัปโหลด หรือรูปแบบไฟล์ไม่ถูกต้องครับ')
    st.stop()

dt_start = data[0]['ts']
dt_end = data[-1]['ts']

duration = dt_end - dt_start
hrs = int(duration.total_seconds() // 3600)
mins = int((duration.total_seconds() % 3600) // 60)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(f'<div class="stat-box"><div class="val">{len(data):,}</div><div class="lbl">Records ทั้งหมด</div></div>', unsafe_allow_html=True)
with col2:
    st.markdown(f'<div class="stat-box"><div class="val">{len(uploaded)}</div><div class="lbl">ไฟล์ .DAD</div></div>', unsafe_allow_html=True)
with col3:
    st.markdown(f'<div class="stat-box"><div class="val">{hrs}h {mins}m</div><div class="lbl">ระยะเวลาข้อมูลที่มี</div></div>', unsafe_allow_html=True)
with col4:
    days_with_data = len(set(r['ts'].strftime('%Y/%m/%d') for r in data))
    st.markdown(f'<div class="stat-box"><div class="val">{days_with_data}</div><div class="lbl">วันที่มีข้อมูล</div></div>', unsafe_allow_html=True)

st.markdown('<div style="margin-top:16px"></div>', unsafe_allow_html=True)

tab_chart, tab_summary, tab_download = st.tabs([
    '📈 กราฟ (Chart)', '📊 ข้อมูลสรุป (Summary)', '⬇️ ดาวน์โหลด Excel'
])

with tab_chart:
    st.markdown('---')
    
    # 1. สร้างพื้นที่เฉพาะให้แต่ละกราฟพร้อม Multiselect ที่เปลี่ยนสีให้ดูดีแล้ว
    sel_top = st.multiselect('🌡️ เลือกแสดง Top Zones', all_names_top, default=all_names_top, key="ms_top")
    fig_top = render_single_chart(data, sel_top, ALL_CH[:7], COLORS_TOP, 'Top Zones', [400, 650], '°C')
    if fig_top: st.plotly_chart(fig_top, use_container_width=True, theme="streamlit")
    st.markdown('<br>', unsafe_allow_html=True)

    sel_bot = st.multiselect('🌡️ เลือกแสดง Bottom Zones', all_names_bot, default=all_names_bot, key="ms_bot")
    fig_bot = render_single_chart(data, sel_bot, ALL_CH[7:14], COLORS_BOT, 'Bottom Zones', [400, 650], '°C')
    if fig_bot: st.plotly_chart(fig_bot, use_container_width=True, theme="streamlit")
    st.markdown('<br>', unsafe_allow_html=True)

    sel_dry = st.multiselect('🔥 เลือกแสดง Dryer Zones', all_names_dry, default=all_names_dry, key="ms_dry")
    fig_dry = render_single_chart(data, sel_dry, ALL_CH[15:17], COLORS_DRY, 'Dryer Zones', [150, 400], '°C')
    if fig_dry: st.plotly_chart(fig_dry, use_container_width=True, theme="streamlit")
    st.markdown('<br>', unsafe_allow_html=True)
    
    sel_o2 = st.multiselect('🫧 เลือกแสดง O2 PPM', all_names_o2, default=all_names_o2, key="ms_o2")
    fig_o2 = render_single_chart(data, sel_o2, [ALL_CH[14], ALL_CH[17]], COLORS_O2, 'O2 PPM', [0, 150], 'ppm')
    if fig_o2: st.plotly_chart(fig_o2, use_container_width=True, theme="streamlit")

with tab_summary:
    st.markdown(f'**ข้อมูลวันที่:** {dt_start.strftime("%d/%m/%Y %H:%M")} → {dt_end.strftime("%d/%m/%Y %H:%M")}')
    st.markdown(f'**ค่าที่กำลังแสดงผล:** {value_mode}')
    st.markdown('---')

    all_sel_names = sel_top + sel_bot + sel_dry + sel_o2
    sel_channels = [(ch,name,unit) for ch,name,unit in ALL_CH if name in all_sel_names]

    rows_summary = []
    for ch,name,unit in sel_channels:
        vals = [r[ch] for r in data if r[ch] is not None]
        if vals:
            rows_summary.append({
                'Zone / Channel': name,
                'Channel': ch,
                'Unit': unit,
                'Avg': round(sum(vals)/len(vals),1),
                'Min': round(min(vals),1),
                'Max': round(max(vals),1),
            })
    if rows_summary:
        df = pd.DataFrame(rows_summary)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning('ยังไม่ได้เลือก Zone หรือไม่มีข้อมูลตัวเลขในโซนที่เลือกครับ')

with tab_download:
    st.markdown('### ⬇️ สร้างและดาวน์โหลดรายงานแบบ Excel')
    all_sel_names = sel_top + sel_bot + sel_dry + sel_o2
    
    if not all_sel_names:
        st.warning('กรุณาเลือก Zone ด้านบนก่อนดาวน์โหลดครับ')
    else:
        if st.button('📊 สร้างไฟล์ Excel', type='primary'):
            with st.spinner('⏳ ระบบกำลังสร้างไฟล์ Excel ให้คุณ...'):
                excel_data = build_excel(data, dt_start, dt_end, all_sel_names)
                fname = f'DAD Report {dt_start.strftime("%d-%m-%Y %H.%M")} to {dt_end.strftime("%d-%m-%Y %H.%M")}.xlsx'
                
                st.download_button(
                    label='⬇️ ดาวน์โหลดรายงาน (Download Excel)',
                    data=excel_data,
                    file_name=fname,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type='primary',
                    use_container_width=True,
                )
                st.success(f'✅ สร้างไฟล์สำเร็จ! สามารถกดปุ่มดาวน์โหลดด้านบนได้เลย')
