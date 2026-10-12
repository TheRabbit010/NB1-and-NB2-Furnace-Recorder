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
    page_title='DAD Report — Valeo NB2',
    page_icon='🏭',
    layout='wide',
    initial_sidebar_state='collapsed',
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
        background: white; border: 1px solid #E0E7EF;
        border-radius: 10px; padding: 14px 18px; text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .stat-box .val { font-size: 1.5rem; font-weight: 700; color: #1F4E79; }
    .stat-box .lbl { font-size: 0.78rem; color: #666; margin-top: 2px; }
    .ok-box   { background:#E2EFDA; border-radius:8px; padding:10px 14px; color:#375623; }
    .warn-box { background:#FFF2CC; border-radius:8px; padding:10px 14px; color:#7D6608; }
    .err-box  { background:#FCE4D6; border-radius:8px; padding:10px 14px; color:#843C0C; }
    
    /* ซ่อน Header และ Toolbar ดั้งเดิม */
    header[data-testid="stHeader"] { background-color: transparent !important; }
    [data-testid="stToolbar"], #MainMenu, .stAppDeployButton { display: none !important; } 
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

COLORS_TOP = ['#378ADD','#D85A30','#1D9E75','#7F77DD','#BA7517','#D4537E','#639922']
COLORS_BOT = ['#53B8E0','#E2864A','#2CB8A0','#9F77CC','#D0913A','#A85090','#5E9C20']
COLORS_O2  = {'CH015':'#2E75B6','CH018':'#C55A11','CH016':'#E36C09','CH017':'#974706'}

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
    green_fill= PatternFill('solid', fgColor='E2EFDA')
    blue_fill = PatternFill('solid', fgColor='DEEAF1')
    red_fill  = PatternFill('solid', fgColor='FCE4D6')
    o2_fill   = PatternFill('solid', fgColor='FFF2CC')
    dryer_fill= PatternFill('solid', fgColor='F0FBF4')
    DAY_FILLS = ['EBF3FB','FFF9EC','F0FBF4','FDF2F8','F5F5F5']

    DATES = sorted(set(r['ts'].strftime('%Y/%m/%d') for r in data))
    date_fill_map = {d: PatternFill('solid', fgColor=DAY_FILLS[i%len(DAY_FILLS)]) for i,d in enumerate(DATES)}
    daily = defaultdict(lambda: {ch:[] for ch,_,_ in ALL_CH})
    
    for r in data:
        d = r['ts'].strftime('%Y/%m/%d')
        for ch,_,_ in ALL_CH:
            if r[ch] is not None: daily[d][ch].append(r[ch])

    def sfill(ch):
        if ch in ('CH015','CH018'): return o2_fill
        if ch in ('CH016','CH017'): return dryer_fill
        idx = int(ch[2:])-1
        return PatternFill('solid', fgColor='EBF3FB') if idx%2==0 else PatternFill('solid', fgColor='FFFFFF')

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

    # Raw Data
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
        bg=date_fill_map.get(dp,PatternFill('solid',fgColor='FFFFFF'))
        dcell(ws,1,ri,ts.strftime('%Y/%m/%d %H:%M:%S'),bg)
        dcell(ws,2,ri,dp,bg); dcell(ws,3,ri,tp,bg)
        for i,(ch,_,_) in enumerate(sel_ch):
            c=ws.cell(ri,4+i,rec[ch]); c.fill=bg; c.font=norm_font
            c.border=border; c.alignment=Alignment(horizontal='right'); c.number_format='0.0'
            
    ws.column_dimensions['A'].width=22; ws.column_dimensions['B'].width=13; ws.column_dimensions['C'].width=10
    for i in range(len(sel_ch)): ws.column_dimensions[get_column_letter(4+i)].width=14
    ws.freeze_panes='A3'; ws.auto_filter.ref=f'A2:{get_column_letter(3+len(sel_ch))}{n+2}'

    # Summary
    ws2=wb.create_sheet('Summary Stats')
    ws2['A1']=f'Summary — {report_name}'
    ws2.merge_cells('A1:H1'); ws2['A1'].font=Font(bold=True,name='Arial',size=13,color='1F4E79')
    ws2['A2']=f'{n:,} records | {dt_start.strftime("%d/%m/%Y %H:%M")} – {dt_end.strftime("%d/%m/%Y %H:%M")}'
    ws2['A2'].font=Font(italic=True,name='Arial',size=9,color='595959')
    for ci,h in enumerate(['Zone / Channel','Channel','Unit','Avg','Min','Max','Range','Std Dev'],1):
        hcell(ws2,ci,4,h)
        
    for zi,(ch,name,unit) in enumerate(sel_ch):
        r=5+zi; col_l=get_column_letter(4+list(ch2 for ch2,_,_ in ALL_CH).index(ch))
        fill=sfill(ch)
        dcell(ws2,1,r,name,fill,bold=True); dcell(ws2,2,r,ch,fill); dcell(ws2,3,r,unit,fill)
        raw_col = get_column_letter(4 + next(i for i,(c,_,_) in enumerate(sel_ch) if c==ch))
        for ci2,formula,f2,fmt in [
            (4,f"=ROUND(AVERAGE('Raw Data'!{raw_col}{ds}:{raw_col}{de}),1)",green_fill,'0.0'),
            (5,f"=ROUND(MIN('Raw Data'!{raw_col}{ds}:{raw_col}{de}),1)",blue_fill,'0.0'),
            (6,f"=ROUND(MAX('Raw Data'!{raw_col}{ds}:{raw_col}{de}),1)",red_fill,'0.0'),
            (7,f"=F{r}-E{r}",PatternFill('solid',fgColor='FFC7CE'),'0.0'),
            (8,f"=ROUND(STDEV('Raw Data'!{raw_col}{ds}:{raw_col}{de}),2)",PatternFill('solid',fgColor='FFFFFF'),'0.00'),
        ]:
            c=ws2.cell(r,ci2,formula); c.fill=f2; c.font=norm_font; c.border=border
            c.alignment=Alignment(horizontal='right'); c.number_format=fmt
    for col,w in zip('ABCDEFGH',[24,10,8,12,12,12,12,10]): ws2.column_dimensions[col].width=w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── Chart Builder ─────────────────────────────────────────────────────────────
def build_chart(data, zones_top, zones_bot, zones_o2):
    rows = 3 if any([zones_top, zones_bot, zones_o2]) else 1
    active_rows = [bool(zones_top), bool(zones_bot), bool(zones_o2)]
    n_rows = sum(active_rows)
    if n_rows == 0: return None

    row_heights = []
    if zones_top: row_heights.append(0.4 if n_rows==1 else 0.35)
    if zones_bot: row_heights.append(0.4 if n_rows==1 else 0.35)
    if zones_o2:  row_heights.append(0.3 if n_rows==1 else 0.25)

    titles = []
    if zones_top: titles.append('🌡️ Top Zones')
    if zones_bot: titles.append('🌡️ Bottom Zones')
    if zones_o2:  titles.append('🫧 O2 & Dryer')

    fig = make_subplots(
        rows=n_rows, cols=1,
        shared_xaxes=True,
        subplot_titles=titles,
        vertical_spacing=0.06,
        row_heights=row_heights,
    )

    def add_traces(ch_list, colors, group, row):
        ts = [r['ts'] for r in data]
        for i,(ch,name,unit) in enumerate(ch_list):
            color = colors[i] if i < len(colors) else COLORS_O2.get(ch,'#888')
            fig.add_trace(go.Scatter(
                x=ts, y=[r[ch] for r in data],
                name=name, mode='lines',
                line=dict(color=color, width=1.5),
                legendgroup=group,
                legendgrouptitle_text=group.title() if i==0 else None,
                hovertemplate=f'<b>{name}</b>: %{{y:.1f}} {unit}<extra></extra>',
            ), row=row, col=1)

    row_num = 1
    if zones_top:
        sel = [(ch,n,u) for ch,n,u in ALL_CH[:7] if n in zones_top]
        colors = [COLORS_TOP[i] for i,(_,n,_) in enumerate(ALL_CH[:7]) if n in zones_top]
        add_traces(sel, colors, 'Top Zones', row_num); row_num += 1
    if zones_bot:
        sel = [(ch,n,u) for ch,n,u in ALL_CH[7:14] if n in zones_bot]
        colors = [COLORS_BOT[i] for i,(_,n,_) in enumerate(ALL_CH[7:14]) if n in zones_bot]
        add_traces(sel, colors, 'Bottom Zones', row_num); row_num += 1
    if zones_o2:
        sel = [(ch,n,u) for ch,n,u in ALL_CH[14:] if n in zones_o2]
        colors = [COLORS_O2.get(ch,'#888') for ch,_,_ in sel]
        add_traces(sel, colors, 'O2 & Dryer', row_num)

    fig.update_layout(
        height=max(400, n_rows * 300),
        template='plotly_white',
        hovermode='x unified',
        hoverlabel=dict(bgcolor='white', font_size=11, namelength=30),
        legend=dict(
            groupclick='toggleitem', bgcolor='rgba(255,255,255,0.95)',
            bordercolor='#C0C8D0', borderwidth=1, font=dict(size=11), tracegroupgap=10,
        ),
        margin=dict(l=60, r=20, t=50, b=20),
        plot_bgcolor='white',
    )
    fig.update_xaxes(
        tickformat='%d/%m\n%H:%M',
        dtick=1*3600*1000, 
        gridcolor='rgba(200,200,200,0.25)',
        tickfont=dict(size=10),
        showticklabels=True,
    )
    for r in range(1, n_rows+1):
        fig.update_yaxes(gridcolor='rgba(200,200,200,0.25)', autorange=True, row=r, col=1)

    return fig

# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════
ICON_B64 = "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
