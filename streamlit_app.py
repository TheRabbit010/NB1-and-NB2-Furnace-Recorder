import streamlit as st
import struct
import io
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title='YOKOGAWA Recorder NB1 & NB2 CAB Furnace',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='expanded',
)

# ── Initialize Session State ──────────────────────────────────────────────────
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F4E79, #2E75B6);
        color: white; padding: 24px 32px; border-radius: 12px;
        margin-bottom: 24px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; font-weight: 600; letter-spacing: 0.5px; }
    .main-header p  { margin: 6px 0 0; opacity: 0.9; font-size: 1rem; }
    .stat-box {
        border: 1px solid var(--secondary-background-color);
        background-color: var(--secondary-background-color);
        border-radius: 10px; padding: 14px 18px; text-align: center;
    }
    .stat-box .val { font-size: 1.5rem; font-weight: 700; color: #2E75B6; }
    .stat-box .lbl { font-size: 0.78rem; opacity: 0.8; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

# ── Constants & Configuration ─────────────────────────────────────────────────
# NB1 Channels
CH_NB1 = [
    ('CH001', 'Top Z#1',       '°C'),
    ('CH002', 'Top Z#2',       '°C'),
    ('CH003', 'Top Z#3',       '°C'),
    ('CH004', 'Top Z#4',       '°C'),
    ('CH005', 'Top Z#5',       '°C'),
    ('CH006', 'Top Z#6',       '°C'),
    ('CH007', 'Top Z#7',       '°C'),
    ('CH008', 'Bottom Z#1',    '°C'),
    ('CH009', 'Bottom Z#2',    '°C'),
    ('CH010', 'Bottom Z#3',    '°C'),
    ('CH011', 'Bottom Z#4',    '°C'),
    ('CH012', 'Bottom Z#5',    '°C'),
    ('CH013', 'Bottom Z#6',    '°C'),
    ('CH014', 'Bottom Z#7',    '°C'),
    ('CH015', 'EXIT O2',       'ppm'),
    ('CH019', 'ENTRANCE O2',   'ppm'),
    ('CH016', 'Dryer #1',      '°C'),
    ('CH017', 'Dryer #2',      '°C'),
    ('CH018', 'N2 .1',         ''),
    ('CH020', 'DEW POINT',     'Cdp'),
]

# NB2 Channels
CH_NB2 = [
    ('CH001', 'Top Z#1',       '°C'),
    ('CH002', 'Top Z#2',       '°C'),
    ('CH003', 'Top Z#3',       '°C'),
    ('CH004', 'Top Z#4',       '°C'),
    ('CH005', 'Top Z#5',       '°C'),
    ('CH006', 'Top Z#6',       '°C'),
    ('CH007', 'Top Z#7',       '°C'),
    ('CH008', 'Bottom Z#1',    '°C'),
    ('CH009', 'Bottom Z#2',    '°C'),
    ('CH010', 'Bottom Z#3',    '°C'),
    ('CH011', 'Bottom Z#4',    '°C'),
    ('CH012', 'Bottom Z#5',    '°C'),
    ('CH013', 'Bottom Z#6',    '°C'),
    ('CH014', 'Bottom Z#7',    '°C'),
    ('CH015', 'O2 EXIT',       'ppm'),
    ('CH018', 'O2 ENTRANCE',   'ppm'),
    ('CH016', 'Dryer #1',      '°C'),
    ('CH017', 'Dryer #2',      '°C'),
]

COLORS_TOP = ['#FF0000', '#008000', '#0000FF', '#8A2BE2', '#A52A2A', '#FFA500', '#9ACD32'] 
COLORS_BOT = ['#ADD8E6', '#FF69B4', '#808080', '#FF0000', '#008000', '#0000FF', '#8A2BE2']
COLORS_DRY = {'CH016':'#FFA500', 'CH017':'#9ACD32'} 
COLORS_O2_NB1 = {'CH015':'#A52A2A', 'CH019':'#ADD8E6'}
COLORS_O2_NB2 = {'CH015':'#A52A2A', 'CH018':'#ADD8E6'}
COLORS_N2 = {'CH018': '#2E75B6'}
COLORS_DEW = {'CH020': '#1D9E75'}

def check_machine_type(filenames):
    has_nb1 = False
    has_nb2 = False
    for f in filenames:
        if 'DATA' in f.upper():
            has_nb2 = True
        else:
            has_nb1 = True
            
    if has_nb1 and has_nb2:
        return 'MIXED', []
    elif has_nb2:
        return 'NB2', CH_NB2
    else:
        return 'NB1', CH_NB1

# ── Auto-Detect BPC & Parser For Binary .DAD ─────────────────────────────────
def find_dad_params(raw):
    max_search = min(50000, len(raw) - 1000)
    for offset in range(10000, max_search):
        yr, mo, dy, hr, mn, sc = raw[offset:offset+6]
        if 0 <= yr <= 99 and 1 <= mo <= 12 and 1 <= dy <= 31 and 0 <= hr <= 23 and 0 <= mn <= 59 and 0 <= sc <= 59:
            for rs in range(24, 1024, 4): 
                if offset + rs*2 + 6 < len(raw):
                    yr2, mo2, dy2, hr2, mn2, sc2 = raw[offset+rs : offset+rs+6]
                    yr3, mo3, dy3, hr3, mn3, sc3 = raw[offset+rs*2 : offset+rs*2+6]
                    
                    def is_valid_date(y, m, d, h, min_, s):
                        return 0 <= y <= 99 and 1 <= m <= 12 and 1 <= d <= 31 and 0 <= h <= 23 and 0 <= min_ <= 59 and 0 <= s <= 59
                        
                    if is_valid_date(yr2, mo2, dy2, hr2, mn2, sc2) and is_valid_date(yr3, mo3, dy3, hr3, mn3, sc3):
                        return offset, rs
    return 15008, 88 

def parse_dad(raw: bytes, active_channels) -> list:
    records = []
    offset, record_size = find_dad_params(raw)
    
    bpc = 4 if record_size >= 70 else 2
    
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
        except ValueError: continue
            
        rec = {'ts': ts}
        for ch, _, unit in active_channels:
            ch_idx = int(ch.replace('CH', '')) - 1
            data_pos = base + 8 + (ch_idx * bpc)
            
            if data_pos + 2 > base + record_size: 
                rec[f'{ch}_min'] = rec[f'{ch}_max'] = rec[f'{ch}_avg'] = rec[ch] = None
                continue
            
            v1 = struct.unpack_from('>h', raw, data_pos)[0]
            if bpc == 4 and data_pos + 4 <= base + record_size:
                v2 = struct.unpack_from('>h', raw, data_pos + 2)[0]
                min_v = min(v1, v2)
                max_v = max(v1, v2)
            else:
                min_v = max_v = v1
            
            divisor = 1.0 if ch == 'CH018' else 10.0
            
            if v1 not in (-32768, -32767, 32767, 32766) and (-30000 < v1 < 30000):
                min_val = min_v / divisor
                max_val = max_v / divisor
                if -1000.0 <= min_val <= 10000.0 and -1000.0 <= max_val <= 10000.0:
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
def build_excel(data, dt_start, dt_end, active_channels, mode_str, machine_type):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Data Viewer'

    ws['A1'] = "YOKOGAWA Recorder Data Analysis Report"
    ws['A1'].font = Font(bold=True, size=12, color='1F4E79')
    ws['A3'], ws['B3'] = "Machine Type", machine_type
    ws['A4'], ws['B4'] = "Data Count", len(data)
    ws['A5'], ws['B5'] = "Start Time", dt_start.strftime("%Y/%m/%d %H:%M:%S")
    ws['A6'], ws['B6'] = "Stop Time", dt_end.strftime("%Y/%m/%d %H:%M:%S")
    ws['A7'], ws['B7'] = "Value Mode", mode_str
    
    hdr_fill = PatternFill('solid', fgColor='F0F4F8')
    hdr_font = Font(bold=True, size=9)
    norm_font = Font(name='Arial', size=9)
    thin = Side(border_style='thin', color='D0D0D0')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.cell(row=9, column=1, value="Ch.").font = hdr_font
    ws.cell(row=10, column=1, value="Tag").font = hdr_font
    ws.cell(row=11, column=1, value="Unit").font = hdr_font
    
    c_date = ws.cell(row=12, column=1, value="Date")
    c_date.font, c_date.border, c_date.fill = hdr_font, border, hdr_fill
    c_time = ws.cell(row=12, column=2, value="Time")
    c_time.font, c_time.border, c_time.fill = hdr_font, border, hdr_fill

    for i, (ch, name, unit) in enumerate(active_channels):
        col = i + 3
        ws.cell(row=9, column=col, value=ch).font = norm_font
        ws.cell(row=10, column=col, value=name).font = norm_font
        ws.cell(row=11, column=col, value=unit).font = norm_font
        c = ws.cell(row=12, column=col, value=mode_str)
        c.font, c.border, c.fill = hdr_font, border, hdr_fill
        c.alignment = Alignment(horizontal='right')

    for ri, rec in enumerate(data, start=13):
        ts = rec['ts']
        ws.cell(row=ri, column=1, value=ts.strftime('%Y/%m/%d')).font = norm_font
        ws.cell(row=ri, column=2, value=ts.strftime('%H:%M:%S')).font = norm_font
        for i, (ch, _, _) in enumerate(active_channels):
            c = ws.cell(row=ri, column=i+3, value=rec[ch])
            c.font, c.number_format = norm_font, '0.0'

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 10
    for i in range(len(active_channels)):
        ws.column_dimensions[get_column_letter(i+3)].width = 11
        
    ws.freeze_panes = 'A13'
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── Chart Builder ─────────────────────────────────────────────────────────────
def render_single_chart(data, ch_list, colors_ref, title, y_range, y_title, hline=None):
    if not ch_list: return None
    fig = go.Figure()
    ts = [r['ts'] for r in data]
    
    for i, (ch, name, unit) in enumerate(ch_list):
        color = colors_ref[i] if isinstance(colors_ref, list) else colors_ref.get(ch, '#888')
        fig.add_trace(go.Scatter(
            x=ts, y=[r[ch] for r in data],
            name=name, mode='lines',
            line=dict(color=color, width=1.0), 
            hovertemplate=f'<b>{name}</b>: %{{y:.1f}} {unit}<extra></extra>' if unit else f'<b>{name}</b>: %{{y:.1f}}<extra></extra>',
        ))

    if hline is not None:
        fig.add_hline(
            y=hline, line_dash="dash", line_color="#FF4B4B", line_width=1.5,
            annotation_text=f"Standard {hline} {y_title}", 
            annotation_position="bottom right",
            annotation_font=dict(color="#FF4B4B", size=12)
        )

    fig.update_layout(
        template='plotly_dark', # ตั้งเป็นโหมด Dark โดยอัตโนมัติ
        title=dict(text=title, font=dict(size=18, weight="bold")),
        height=450, hovermode='x unified', hoverlabel=dict(font_size=12, namelength=30),
        legend=dict(
            title="Channels", orientation="v", yanchor="middle", y=0.5, 
            xanchor="left", x=1.02, font=dict(size=12), itemsizing='constant'
        ),
        margin=dict(l=60, r=150, t=60, b=40),
    )
    
    if y_range: fig.update_yaxes(range=y_range, title_text=y_title, title_font=dict(size=14))
    else: fig.update_yaxes(title_text=y_title, title_font=dict(size=14))
    fig.update_xaxes(tickformat='%d/%m\n%H:%M', tickfont=dict(size=11), showticklabels=True)
    return fig

def render_dual_axis_chart(data, ch_left, ch_right, colors_left, colors_right, title, y1_range, y1_title, y2_title, hline_left=None, y2_range=None):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    ts = [r['ts'] for r in data]
    
    for i, (ch, name, unit) in enumerate(ch_left):
        color = colors_left[i] if isinstance(colors_left, list) else colors_left.get(ch, '#888')
        fig.add_trace(go.Scatter(
            x=ts, y=[r[ch] for r in data], name=name, mode='lines', 
            line=dict(color=color, width=1.0),
            hovertemplate=f'<b>{name}</b>: %{{y:.1f}} {unit}<extra></extra>' if unit else f'<b>{name}</b>: %{{y:.1f}}<extra></extra>',
        ), secondary_y=False)

    for i, (ch, name, unit) in enumerate(ch_right):
        color = colors_right[i] if isinstance(colors_right, list) else colors_right.get(ch, '#888')
        fig.add_trace(go.Scatter(
            x=ts, y=[r[ch] for r in data], name=name, mode='lines', 
            line=dict(color=color, width=1.0, dash='dot'), 
            hovertemplate=f'<b>{name}</b>: %{{y:.0f}} {unit}<extra></extra>' if unit else f'<b>{name}</b>: %{{y:.0f}}<extra></extra>',
        ), secondary_y=True)

    if hline_left is not None:
        fig.add_hline(
            y=hline_left, line_dash="dash", line_color="#FF4B4B", line_width=1.5,
            annotation_text=f"Standard {hline_left} {y1_title.split()[0]}", 
            annotation_position="bottom right", annotation_font=dict(color="#FF4B4B", size=12), secondary_y=False
        )

    fig.update_layout(
        template='plotly_dark', # ตั้งเป็นโหมด Dark โดยอัตโนมัติ
        title=dict(text=title, font=dict(size=18, weight="bold")),
        height=450, hovermode='x unified', hoverlabel=dict(font_size=12, namelength=30),
        legend=dict(
            title="Channels", orientation="v", yanchor="middle", y=0.5, 
            xanchor="left", x=1.08, font=dict(size=12), itemsizing='constant'
        ),
        margin=dict(l=60, r=180, t=60, b=40),
    )
    
    fig.update_yaxes(range=y1_range, title_text=y1_title, title_font=dict(size=14), secondary_y=False)
    if y2_range: fig.update_yaxes(range=y2_range, title_text=y2_title, title_font=dict(size=14), secondary_y=True, showgrid=False)
    else: fig.update_yaxes(title_text=y2_title, title_font=dict(size=14), secondary_y=True, showgrid=False)
    fig.update_xaxes(tickformat='%d/%m\n%H:%M', tickfont=dict(size=11), showticklabels=True)
    return fig

# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="main-header">
    <h1>YOKOGAWA Recorder NB1 & NB2 CAB Furnace</h1>
    <p>Furnace Data Analysis Report</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    if st.button("🗑️ ล้างข้อมูล / หยุดการทำงานทันที", type="primary", use_container_width=True):
        st.session_state.uploader_key += 1
        st.cache_data.clear()
        st.rerun()

    st.markdown('<br>', unsafe_allow_html=True)
    st.markdown('### 📂 Upload ไฟล์ .DAD')
    
    uploaded = st.file_uploader(
        'ลากไฟล์ .DAD มาวางที่นี่ (เลือกได้หลายไฟล์)',
        type=['DAD','dad'],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.uploader_key}"
    )
    
    st.markdown("---")
    st.markdown('### ⚙️ ตั้งค่าการอ่านไฟล์และกราฟ')
    value_mode = st.selectbox(
        'เลือกประเภทของค่าที่ต้องการวิเคราะห์:',
        options=['MIN', 'AVG (เฉลี่ย)', 'MAX'],
        index=2 
    )
    mode_key = {'MIN': '_min', 'AVG (เฉลี่ย)': '_avg', 'MAX': '_max'}[value_mode]

if not uploaded:
    st.info('👉 กรุณาอัปโหลดไฟล์ .DAD จากแถบเมนูด้านซ้ายเพื่อเริ่มต้นสร้างรายงานวิเคราะห์ข้อมูล')
    st.stop()

files_data = [(f.name, f.read()) for f in uploaded]
machine_type, active_channels = check_machine_type([f[0] for f in files_data])

if machine_type == 'MIXED':
    st.error('🚨 ตรวจพบการอัปโหลดไฟล์จากทั้งเครื่อง NB1 และ NB2 ปะปนกัน! กรุณากดปุ่ม "ล้างข้อมูล" และอัปโหลดไฟล์จากเครื่องใดเครื่องหนึ่งเท่านั้นครับ')
    st.stop()

@st.cache_data(show_spinner=f'⏳ กำลังประมวลผลไฟล์จากเครื่อง {machine_type}...')
def process_dad_files(files_d, mode_k, act_channels):
    all_records = []
    for fname, raw in files_d:
        recs = parse_dad(raw, act_channels)
        
        if len(recs) > 10:
            all_records.extend(recs[3:-2])
        else:
            all_records.extend(recs)
        
    if not all_records: return []
    all_records.sort(key=lambda x: x['ts'])
    
    if len(all_records) > 0:
        streaks = []
        current_streak = [all_records[0]]
        for i in range(1, len(all_records)):
            if (all_records[i]['ts'] - all_records[i-1]['ts']).total_seconds() <= 86400 * 2:
                current_streak.append(all_records[i])
            else:
                streaks.append(current_streak)
                current_streak = [all_records[i]]
        streaks.append(current_streak)
        all_records = max(streaks, key=len)
        
    seen = set()
    unique = []
    for r in all_records:
        k = r['ts']
        if k not in seen:
            seen.add(k)
            new_r = dict(r)
            for ch,_,_ in act_channels:
                new_r[ch] = r.get(f'{ch}{mode_k}')
            unique.append(new_r)
    return unique

data = process_dad_files(tuple((n,d) for n,d in files_data), mode_key, tuple(active_channels))

if len(data) == 0:
    st.error('🚨 ไม่พบข้อมูลในไฟล์ที่อัปโหลด กรุณาตรวจสอบความถูกต้องของไฟล์')
    st.stop()

dt_start = data[0]['ts']
dt_end = data[-1]['ts']
duration = dt_end - dt_start
hrs = int(duration.total_seconds() // 3600)
mins = int((duration.total_seconds() % 3600) // 60)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(f'<div class="stat-box"><div class="val">{machine_type}</div><div class="lbl">Detected Machine</div></div>', unsafe_allow_html=True)
with col2:
    st.markdown(f'<div class="stat-box"><div class="val">{len(data):,}</div><div class="lbl">Total Records</div></div>', unsafe_allow_html=True)
with col3:
    st.markdown(f'<div class="stat-box"><div class="val">{hrs}h {mins}m</div><div class="lbl">Data Duration</div></div>', unsafe_allow_html=True)
with col4:
    st.markdown(f'<div class="stat-box"><div class="val">{value_mode.split()[0]}</div><div class="lbl">Current Mode</div></div>', unsafe_allow_html=True)

st.markdown('<div style="margin-top:24px"></div>', unsafe_allow_html=True)

tab_chart, tab_summary, tab_download = st.tabs([
    '📈 กราฟแสดงผล (Chart)', '📊 ข้อมูลสรุป (Summary)', '⬇️ ดาวน์โหลดรายงาน (Excel)'
])

with tab_chart:
    st.markdown('<br>', unsafe_allow_html=True)
    
    tb_y_range = [540, 650] if machine_type == 'NB1' else [400, 650]
    
    top_chs = [ch for ch in active_channels if ch[0] in ['CH001','CH002','CH003','CH004','CH005','CH006','CH007']]
    fig_top = render_single_chart(data, top_chs, COLORS_TOP, f'Top Zones Temperature ({machine_type})', tb_y_range, '°C')
    # ใช้ theme=None เพื่อให้แสดงผลเป็น template='plotly_dark' อย่างสมบูรณ์โดยไม่ถูก streamlit override กลับ
    if fig_top: st.plotly_chart(fig_top, use_container_width=True, theme=None)
    st.markdown('<br>', unsafe_allow_html=True)

    bot_chs = [ch for ch in active_channels if ch[0] in ['CH008','CH009','CH010','CH011','CH012','CH013','CH014']]
    fig_bot = render_single_chart(data, bot_chs, COLORS_BOT, f'Bottom Zones Temperature ({machine_type})', tb_y_range, '°C')
    if fig_bot: st.plotly_chart(fig_bot, use_container_width=True, theme=None)
    st.markdown('<br>', unsafe_allow_html=True)

    dry_chs = [ch for ch in active_channels if ch[0] in ['CH016','CH017']]
    dry_y_range = [150, 350] if machine_type == 'NB1' else [150, 400]
    fig_dry = render_single_chart(data, dry_chs, COLORS_DRY, f'Dryer Zones Temperature ({machine_type})', dry_y_range, '°C')
    if fig_dry: st.plotly_chart(fig_dry, use_container_width=True, theme=None)
    st.markdown('<br>', unsafe_allow_html=True)
    
    if machine_type == 'NB2':
        o2_chs = [ch for ch in active_channels if ch[0] in ['CH015', 'CH018']]
        fig_o2 = render_single_chart(data, o2_chs, COLORS_O2_NB2, 'O2 Concentration Level', [0, 150], 'ppm', hline=100)
        if fig_o2: st.plotly_chart(fig_o2, use_container_width=True, theme=None)
    else:
        o2_chs = [ch for ch in active_channels if ch[0] in ['CH015', 'CH019']]
        n2_chs = [ch for ch in active_channels if ch[0] == 'CH018']
        fig_o2_n2 = render_dual_axis_chart(
            data, o2_chs, n2_chs, 
            COLORS_O2_NB1, COLORS_N2, 
            'O2 Concentration & N2 Flow Rate', 
            [0, 150], 'O2 (ppm)', '', 
            hline_left=100, y2_range=[0, 5000]
        )
        if fig_o2_n2: st.plotly_chart(fig_o2_n2, use_container_width=True, theme=None)
        st.markdown('<br>', unsafe_allow_html=True)
        
        dew_chs = [ch for ch in active_channels if ch[0] == 'CH020']
        fig_dew = render_single_chart(data, dew_chs, COLORS_DEW, 'Dew Point', [-100, -10], 'Cdp')
        if fig_dew: st.plotly_chart(fig_dew, use_container_width=True, theme=None)

with tab_summary:
    st.markdown(f'**ช่วงเวลาของข้อมูล:** {dt_start.strftime("%d/%m/%Y %H:%M")} → {dt_end.strftime("%d/%m/%Y %H:%M")}')
    st.markdown(f'**เครื่องที่วิเคราะห์:** {machine_type} | **การประมวลผล:** {value_mode}')
    st.markdown('---')

    rows_summary = []
    for ch, name, unit in active_channels:
        vals = [r[ch] for r in data if r[ch] is not None]
        if vals:
            rows_summary.append({
                'Zone / Channel': name,
                'Channel ID': ch,
                'Unit': unit,
                'Avg': round(sum(vals)/len(vals),1),
                'Min': round(min(vals),1),
                'Max': round(max(vals),1),
            })
    if rows_summary:
        df = pd.DataFrame(rows_summary)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning('ไม่มีข้อมูลตัวเลขในไฟล์ที่ประมวลผล')

with tab_download:
    st.markdown(f'### ⬇️ Export Data Report (Excel) - {machine_type}')
    
    if st.button('📊 กดเพื่อสร้างไฟล์ Excel', type='primary'):
        with st.spinner('⏳ ระบบกำลังเตรียมไฟล์ Excel...'):
            mode_str = value_mode.split()[0]
            excel_data = build_excel(data, dt_start, dt_end, active_channels, mode_str, machine_type)
            fname = f'DAD_Report_{machine_type}_{dt_start.strftime("%Y%m%d_%H%M")}.xlsx'
            
            st.download_button(
                label='⬇️ ดาวน์โหลดไฟล์ Excel (Click to Download)',
                data=excel_data,
                file_name=fname,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
                use_container_width=True,
            )
            st.success(f'✅ ไฟล์พร้อมดาวน์โหลดแล้ว!')
