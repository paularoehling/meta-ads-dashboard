#!/usr/bin/env python3
import json, base64, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def fill(c): return PatternFill('solid', fgColor=c.replace('#',''))
def thin(): return Side(style='thin', color='CCCCCC')
def bdr(): return Border(left=thin(),right=thin(),top=thin(),bottom=thin())
def aln(h='left',v='center',wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

def w(ws,row,col,val,bold=False,fg='1E2A38',bg=None,align='left',
      size=10,border=True,wrap=False,italic=False,fmt=None):
    c=ws.cell(row=row,column=col,value=val)
    c.font=Font(name='Arial',bold=bold,color=fg.replace('#',''),size=size,italic=italic)
    if bg: c.fill=fill(bg)
    c.alignment=aln(align,wrap=wrap)
    if border: c.border=bdr()
    if fmt: c.number_format=fmt
    return c

def get_status(tipo, costo):
    if not costo or costo<=0: return '—'
    t=(tipo or '').lower()
    if 'purchase' in t or 'convers' in t or 'venta' in t or 'compra' in t:
        if costo<6000: return '🏆 GANADOR'
        if costo<10000: return '✅ OK'
        if costo<18000: return '🔍 REVISAR'
        return '⛔ PAUSAR'
    elif 'lead' in t or 'formulario' in t:
        if costo<50: return '🏆 GANADOR'
        if costo<90: return '✅ OK'
        if costo<150: return '🔍 REVISAR'
        return '⛔ PAUSAR'
    else:
        if costo<20: return '🏆 GANADOR'
        if costo<40: return '✅ OK'
        if costo<80: return '🔍 REVISAR'
        return '⛔ PAUSAR'

with open('/tmp/excel_data.json','r') as f:
    data = json.load(f)

label_actual   = data.get('label_actual', 'Mes actual')
label_anterior = data.get('label_anterior', 'Mes anterior')
marca_nombre   = data.get('marca_nombre', 'Marca')
ads_actual     = data.get('ads_actual', [])
ads_anterior   = data.get('ads_anterior', [])



# Lookup anterior por nombre de anuncio
ant_map = {a.get('anuncio',''): a for a in ads_anterior}

wb = Workbook()
ws = wb.active
ws.title = 'RESULTADOS'

# ── HEADER ────────────────────────────────────────────────────────────────────
ws.merge_cells('B1:T1')
w(ws,1,2,f'INFORME META ADS — {marca_nombre.upper()} — {label_actual.upper()}',
  bold=True,fg='FFFFFF',bg='1E2A38',align='center',size=13)
ws.row_dimensions[1].height=34

ws.merge_cells('B2:T2')
w(ws,2,2,f'{label_actual} vs {label_anterior} · Fuente: Meta Ads API · Agrupado por campaña',
  fg='AAAAAA',bg='1E2A38',align='center',size=10)
ws.row_dimensions[2].height=18

# ── HEADERS COLUMNAS ─────────────────────────────────────────────────────────
headers=['CAMPAÑA / ANUNCIO','Alcance','Impresiones','Frecuencia',
         'Resultados',f'Resultado ant.\n({label_anterior})','%',
         'Gasto (CLP)','Costo/Result.','CTR%','CPM',
         'Hook Rate%','Hold Rate 50%','Hold Rate 75%','Hold Rate 100%',
         'ROAS','Clicks sitio','STATUS']
for i,h in enumerate(headers):
    w(ws,5,i+2,h,bold=True,fg='FFFFFF',bg='1E2A38',align='center',size=10,wrap=True)
ws.row_dimensions[5].height=45

col_widths=[35,12,13,10,12,16,8,14,12,8,10,10,13,13,13,8,11,14]
for i,cw in enumerate(col_widths):
    ws.column_dimensions[get_column_letter(i+2)].width=cw

# ── DATOS ─────────────────────────────────────────────────────────────────────
row=7
camps={}
for ad in ads_actual:
    cn=ad.get('campaña','Sin campaña')
    if cn not in camps: camps[cn]=[]
    camps[cn].append(ad)

for camp_name, ads in camps.items():
    # Header campaña
    ws.merge_cells(f'B{row}:T{row}')
    w(ws,row,2,f'▶ {camp_name}',bold=True,fg='FFFFFF',bg='34495E',align='left',size=10)
    ws.row_dimensions[row].height=22
    row+=1

    tot_res=tot_res_ant=tot_gasto=0

    for ad in ads:
        nombre   = ad.get('anuncio','')
        tipo     = ad.get('tipo','')
        res      = float(ad.get('resultados',0) or 0)
        gasto    = float(ad.get('gasto',0) or 0)
        costo    = float(ad.get('costo',0) or 0)
        ctr      = float(ad.get('ctr',0) or 0)
        alcance  = float(ad.get('alcance',0) or 0)
        impr     = float(ad.get('impresiones',0) or 0)
        freq     = float(ad.get('frecuencia',0) or 0)
        hook     = float(ad.get('hook',0) or 0)
        hold50   = float(ad.get('hold50',0) or 0)
        hold75   = float(ad.get('hold75',0) or 0)
        hold100  = float(ad.get('hold100',0) or 0)
        roas     = float(ad.get('roas',0) or 0)
        clicks   = float(ad.get('clicks_sitio',0) or 0)
        cpm      = float(ad.get('cpm',0) or 0)

        # Datos mes anterior
        ad_ant   = ant_map.get(nombre,{})
        res_ant  = float(ad_ant.get('resultados',0) or 0)
        pct      = ((res-res_ant)/res_ant*100) if res_ant>0 else None
        status   = get_status(tipo, costo)

        # Color fila
        bg='FFFFFF'
        if 'PAUSAR' in status: bg='FDEDEC'
        elif 'GANADOR' in status: bg='EAFAF1'
        elif 'REVISAR' in status: bg='FEF9E7'

        w(ws,row,2,nombre,fg='1E2A38',bg=bg)
        w(ws,row,3,int(alcance) if alcance else '',bg=bg,align='right')
        w(ws,row,4,int(impr) if impr else '',bg=bg,align='right')
        w(ws,row,5,round(freq,1) if freq else '',bg=bg,align='right')
        w(ws,row,6,int(res) if res else '',bold=True,bg=bg,align='right')
        w(ws,row,7,int(res_ant) if res_ant else '—',
          italic=bool(res_ant),fg='888888' if not res_ant else '1E2A38',bg=bg,align='right')
        if pct is not None:
            pct_fg='27AE60' if pct>0 else 'E74C3C'
            w(ws,row,8,('+' if pct>0 else '')+f'{pct:.0f}%',bold=True,fg=pct_fg,bg=bg,align='center')
        else:
            w(ws,row,8,'—',fg='AAAAAA',bg=bg,align='center')
        w(ws,row,9,int(gasto) if gasto else '',bg=bg,align='right',fmt='#,##0')
        w(ws,row,10,int(costo) if costo else '',bg=bg,align='right',fmt='#,##0')
        w(ws,row,11,round(ctr,2) if ctr else '',bg=bg,align='right')
        w(ws,row,12,int(cpm) if cpm else '',bg=bg,align='right')
        hr_fg='27AE60' if hook>=25 else ('E74C3C' if 0<hook<15 else '1E2A38')
        w(ws,row,13,round(hook,1) if hook else '',bold=hook>=25,fg=hr_fg,bg=bg,align='right')
        w(ws,row,14,round(hold50,1) if hold50 else '',bg=bg,align='right')
        w(ws,row,15,round(hold75,1) if hold75 else '',bg=bg,align='right')
        w(ws,row,16,round(hold100,1) if hold100 else '',bg=bg,align='right')
        w(ws,row,17,round(roas,2) if roas else '',bg=bg,align='right')
        w(ws,row,18,int(clicks) if clicks else '',bg=bg,align='right')
        status_fg='27AE60' if 'OK' in status or 'GANADOR' in status else ('E74C3C' if 'PAUSAR' in status else 'E67E22')
        w(ws,row,19,status,bold=True,bg=bg,align='center',fg=status_fg)
        ws.row_dimensions[row].height=18

        tot_res+=res; tot_res_ant+=res_ant; tot_gasto+=gasto
        row+=1

    # Subtotal
    ws.merge_cells(f'B{row}:E{row}')
    w(ws,row,2,f'SUBTOTAL  {camp_name}',bold=True,fg='FFFFFF',bg='2E4057')
    w(ws,row,6,int(tot_res),bold=True,fg='FFFFFF',bg='2E4057',align='right')
    if tot_res_ant>0:
        w(ws,row,7,int(tot_res_ant),bold=True,fg='FFFFFF',bg='2E4057',align='right')
        p=((tot_res-tot_res_ant)/tot_res_ant*100)
        w(ws,row,8,('+' if p>0 else '')+f'{p:.0f}%',bold=True,fg='FFFFFF',bg='2E4057',align='center')
    w(ws,row,9,int(tot_gasto),bold=True,fg='FFFFFF',bg='2E4057',align='right',fmt='#,##0')
    ws.row_dimensions[row].height=20
    row+=2

ws.freeze_panes='B6'
ws.sheet_view.zoomScale=90

buf=io.BytesIO()
wb.save(buf)
buf.seek(0)
print(base64.b64encode(buf.read()).decode('utf-8'))
