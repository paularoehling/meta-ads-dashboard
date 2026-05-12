#!/usr/bin/env python3
"""
gen_excel.py — Meta Ads Report Generator
Lee /tmp/excel_data.json, genera el Excel y lo imprime en base64
"""
import json, sys, base64, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def fill(c): return PatternFill('solid', fgColor=c.replace('#',''))
def thin(): return Side(style='thin', color='CCCCCC')
def bdr(): return Border(left=thin(),right=thin(),top=thin(),bottom=thin())
def aln(h='left',v='center',wrap=False): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
def fnt(bold=False,color='1E2A38',size=10,italic=False):
    return Font(name='Arial',bold=bold,color=color.replace('#',''),size=size,italic=italic)

def w(ws,row,col,val,bold=False,fg='1E2A38',bg=None,align='left',
      size=10,border=True,wrap=False,italic=False,fmt=None):
    c=ws.cell(row=row,column=col,value=val)
    c.font=fnt(bold,fg,size,italic)
    if bg: c.fill=fill(bg)
    c.alignment=aln(align,wrap=wrap)
    if border: c.border=bdr()
    if fmt: c.number_format=fmt
    return c

def get_action(actions, action_type):
    if not actions: return 0
    for a in actions:
        if a.get('action_type')==action_type:
            try: return float(a.get('value',0))
            except: return 0
    return 0

def get_status(objetivo, costo):
    if not costo or costo<=0: return '—'
    obj = objetivo.lower() if objetivo else ''
    if 'purchase' in obj or 'convers' in obj or 'venta' in obj or 'compra' in obj:
        if costo < 6000: return '🏆 GANADOR'
        if costo < 10000: return '✅ OK'
        if costo < 18000: return '🔍 REVISAR'
        return '⛔ PAUSAR'
    elif 'lead' in obj or 'formulario' in obj:
        if costo < 50: return '🏆 GANADOR'
        if costo < 90: return '✅ OK'
        if costo < 150: return '🔍 REVISAR'
        return '⛔ PAUSAR'
    elif 'link_click' in obj or 'trafico' in obj or 'tráfico' in obj or 'click' in obj:
        if costo < 20: return '🏆 GANADOR'
        if costo < 40: return '✅ OK'
        if costo < 80: return '🔍 REVISAR'
        return '⛔ PAUSAR'
    elif 'video' in obj or 'perfil' in obj or 'instagram' in obj or 'visita' in obj:
        if costo < 20: return '🏆 GANADOR'
        if costo < 40: return '✅ OK'
        if costo < 80: return '🔍 REVISAR'
        return '⛔ PAUSAR'
    else:
        if costo < 10000: return '✅ OK'
        return '🔍 REVISAR'

# ── LEER DATOS ─────────────────────────────────────────────────────────────────
with open('/tmp/excel_data.json','r') as f:
    data = json.load(f)

marcas = data.get('marcas', [])
mes_actual = data.get('mes_actual', 'Mes actual')
mes_anterior = data.get('mes_anterior', 'Mes anterior')

wb = Workbook()

# ── HOJA RESULTADOS ────────────────────────────────────────────────────────────
ws = wb.active
ws.title = 'RESULTADOS'

# Cabecera principal
ws.merge_cells('B1:T1')
w(ws,1,2,f'INFORME META ADS — {mes_actual.upper()}',bold=True,fg='FFFFFF',bg='1E2A38',align='center',size=14)
ws.row_dimensions[1].height=36

ws.merge_cells('B2:T2')
w(ws,2,2,f'{mes_actual} vs {mes_anterior} · Fuente: Meta Ads API · Agrupado por campaña',
  fg='AAAAAA',bg='1E2A38',align='center',size=10)
ws.row_dimensions[2].height=20

# Benchmarks fila 5
bench_labels=['','>1%','','>25%','>12%','>8%','>3%','>2x']
for i,lbl in enumerate(bench_labels):
    col=13+i
    if lbl and col<=20:
        w(ws,5,col,lbl,bold=True,fg='666666',bg='F8F8F8',align='center',size=9,border=False)

# Headers fila 7
headers=['CAMPAÑA / ANUNCIO','Alcance','Impresiones','Frecuencia','Resultados',
         f'Resultado ant.\n({mes_anterior})','%','Gasto (CLP)','Costo/Result.',
         'CTR%','CPM','Hook\nRate%','Hold Rate 50%','Hold Rate 75%','Hold Rate\n100%',
         'ROAS','Clicks sitio','STATUS']
for i,h in enumerate(headers):
    col=i+2
    w(ws,7,col,h,bold=True,fg='FFFFFF',bg='1E2A38',align='center',size=10,wrap=True)
ws.row_dimensions[7].height=45

# Anchos de columna
col_widths=[35,12,13,10,12,16,8,14,12,8,10,10,13,13,13,8,11,14]
for i,cw in enumerate(col_widths):
    ws.column_dimensions[get_column_letter(i+2)].width=cw

row=9
for marca in marcas:
    # Fila de marca
    ws.merge_cells(f'B{row}:T{row}')
    w(ws,row,2,marca['nombre'],bold=True,fg='FFFFFF',bg='2C3E50',align='left',size=11)
    ws.row_dimensions[row].height=24
    row+=1

    # Campañas
    camps={}
    for ad in marca.get('ads',[]):
        cn=ad.get('campana','Sin campaña')
        if cn not in camps: camps[cn]=[]
        camps[cn].append(ad)

    for camp_name, ads in camps.items():
        # Header campaña
        ws.merge_cells(f'B{row}:T{row}')
        w(ws,row,2,f'▶ {camp_name}',bold=True,fg='FFFFFF',bg='34495E',align='left',size=10)
        ws.row_dimensions[row].height=22
        row+=1

        tot_res=tot_res_ant=tot_gasto=tot_alc=tot_imp=0
        obj_camp=''

        for ad in ads:
            spend=float(ad.get('spend',0) or 0)
            reach=float(ad.get('reach',0) or 0)
            impr=float(ad.get('impressions',0) or 0)
            freq=float(ad.get('frequency',0) or 0)
            ctr=float(ad.get('ctr',0) or 0)
            cpm=float(ad.get('cpm',0) or 0)

            actions=ad.get('actions',[])
            objetivo=ad.get('objetivo','')
            obj_camp=objetivo

            # Resultado según objetivo
            if 'purchase' in objetivo.lower():
                res=get_action(actions,'purchase')
            elif 'lead' in objetivo.lower():
                res=get_action(actions,'lead')
            elif 'link_click' in objetivo.lower():
                res=get_action(actions,'link_click')
            elif 'video' in objetivo.lower() or 'visita' in objetivo.lower():
                res=get_action(actions,'post_engagement')
            else:
                res=sum(float(a.get('value',0)) for a in actions if a) if actions else 0

            res_ant=float(ad.get('resultado_anterior',0) or 0)
            costo_r=spend/res if res>0 else 0
            status=get_status(objetivo, costo_r)

            # Hook y Hold rates
            video_plays=get_action(ad.get('video_play_actions',[]),'video_view')
            hook_rate=0
            if video_plays>0 and impr>0:
                hook_rate=(video_plays/impr)*100

            p50=float(ad.get('video_p50',0) or 0)
            p75=float(ad.get('video_p75',0) or 0)
            p100=float(ad.get('video_p100',0) or 0)
            hold50=p50/impr*100 if impr>0 and p50>0 else 0
            hold75=p75/impr*100 if impr>0 and p75>0 else 0
            hold100=p100/impr*100 if impr>0 and p100>0 else 0

            roas_list=ad.get('purchase_roas',[])
            roas=float(roas_list[0].get('value',0)) if roas_list else 0
            clicks=get_action(ad.get('outbound_clicks',[]),'outbound_click')

            pct_cambio=((res-res_ant)/res_ant*100) if res_ant>0 else None

            # Color de fila
            bg_row='FFFFFF'
            if status=='⛔ PAUSAR': bg_row='FDEDEC'
            elif status=='🏆 GANADOR': bg_row='EAFAF1'
            elif status=='🔍 REVISAR': bg_row='FEF9E7'

            w(ws,row,2,ad.get('nombre',''),fg='1E2A38',bg=bg_row)
            w(ws,row,3,int(reach) if reach else '',bg=bg_row,align='right')
            w(ws,row,4,int(impr) if impr else '',bg=bg_row,align='right')
            w(ws,row,5,round(freq,1) if freq else '',bg=bg_row,align='right')
            w(ws,row,6,int(res) if res else '',bold=True,bg=bg_row,align='right')

            if res_ant>0:
                w(ws,row,7,int(res_ant),italic=True,fg='888888',bg=bg_row,align='right')
            else:
                w(ws,row,7,'—',fg='AAAAAA',bg=bg_row,align='center')

            if pct_cambio is not None:
                pct_fg='27AE60' if pct_cambio>0 else 'E74C3C'
                pct_str=('+' if pct_cambio>0 else '')+f'{pct_cambio:.0f}%'
                w(ws,row,8,pct_str,bold=True,fg=pct_fg,bg=bg_row,align='center')
            else:
                w(ws,row,8,'—',fg='AAAAAA',bg=bg_row,align='center')

            w(ws,row,9,int(spend) if spend else '',bg=bg_row,align='right',fmt='#,##0')
            w(ws,row,10,int(costo_r) if costo_r else '',bg=bg_row,align='right',fmt='#,##0')
            w(ws,row,11,round(ctr,2) if ctr else '',bg=bg_row,align='right')
            w(ws,row,12,int(cpm) if cpm else '',bg=bg_row,align='right')

            hr_fg='27AE60' if hook_rate>=25 else ('E74C3C' if hook_rate>0 and hook_rate<15 else '1E2A38')
            w(ws,row,13,round(hook_rate,1) if hook_rate else '',bold=hook_rate>=25,fg=hr_fg,bg=bg_row,align='right')
            w(ws,row,14,round(hold50,1) if hold50 else '',bg=bg_row,align='right')
            w(ws,row,15,round(hold75,1) if hold75 else '',bg=bg_row,align='right')
            w(ws,row,16,round(hold100,1) if hold100 else '',bg=bg_row,align='right')
            w(ws,row,17,round(roas,2) if roas else '',bg=bg_row,align='right')
            w(ws,row,18,int(clicks) if clicks else '',bg=bg_row,align='right')
            w(ws,row,19,status,bold=True,bg=bg_row,align='center',
              fg='27AE60' if 'OK' in status or 'GANADOR' in status else ('E74C3C' if 'PAUSAR' in status else 'E67E22'))

            ws.row_dimensions[row].height=18
            tot_res+=res; tot_res_ant+=res_ant; tot_gasto+=spend
            tot_alc+=reach; tot_imp+=impr
            row+=1

        # Subtotal campaña
        ws.merge_cells(f'B{row}:E{row}')
        w(ws,row,2,f'SUBTOTAL  {camp_name}',bold=True,fg='FFFFFF',bg='2E4057',align='left')
        w(ws,row,6,int(tot_res),bold=True,fg='FFFFFF',bg='2E4057',align='right')
        if tot_res_ant>0:
            w(ws,row,7,int(tot_res_ant),bold=True,fg='FFFFFF',bg='2E4057',align='right')
            pct_tot=((tot_res-tot_res_ant)/tot_res_ant*100)
            pct_s=('+' if pct_tot>0 else '')+f'{pct_tot:.0f}%'
            w(ws,row,8,pct_s,bold=True,fg='FFFFFF',bg='2E4057',align='center')
        w(ws,row,9,int(tot_gasto),bold=True,fg='FFFFFF',bg='2E4057',align='right',fmt='#,##0')
        ws.row_dimensions[row].height=20
        row+=1

    row+=1  # espacio entre marcas

# Freeze y zoom
ws.freeze_panes='B8'
ws.sheet_view.zoomScale=90

# ── HOJAS POR MARCA ───────────────────────────────────────────────────────────
for marca in marcas:
    ws_m=wb.create_sheet(title=marca['nombre'][:31])

    ws_m.merge_cells('B1:T1')
    w(ws_m,1,2,f"{marca['nombre'].upper()} — {mes_actual.upper()}",
      bold=True,fg='FFFFFF',bg='1E2A38',align='center',size=13)
    ws_m.row_dimensions[1].height=32

    for i,h in enumerate(headers):
        col=i+2
        w(ws_m,3,col,h,bold=True,fg='FFFFFF',bg='1E2A38',align='center',size=10,wrap=True)
    ws_m.row_dimensions[3].height=45

    for i,cw in enumerate(col_widths):
        ws_m.column_dimensions[get_column_letter(i+2)].width=cw

    mrow=5
    camps_m={}
    for ad in marca.get('ads',[]):
        cn=ad.get('campana','Sin campaña')
        if cn not in camps_m: camps_m[cn]=[]
        camps_m[cn].append(ad)

    for camp_name,ads in camps_m.items():
        ws_m.merge_cells(f'B{mrow}:T{mrow}')
        w(ws_m,mrow,2,f'▶ {camp_name}',bold=True,fg='FFFFFF',bg='34495E',align='left',size=10)
        ws_m.row_dimensions[mrow].height=22
        mrow+=1

        st_res=st_ant=st_gasto=0
        for ad in ads:
            spend=float(ad.get('spend',0) or 0)
            reach=float(ad.get('reach',0) or 0)
            impr=float(ad.get('impressions',0) or 0)
            freq=float(ad.get('frequency',0) or 0)
            ctr=float(ad.get('ctr',0) or 0)
            cpm=float(ad.get('cpm',0) or 0)
            actions=ad.get('actions',[])
            objetivo=ad.get('objetivo','')

            if 'purchase' in objetivo.lower(): res=get_action(actions,'purchase')
            elif 'lead' in objetivo.lower(): res=get_action(actions,'lead')
            elif 'link_click' in objetivo.lower(): res=get_action(actions,'link_click')
            else: res=sum(float(a.get('value',0)) for a in actions if a) if actions else 0

            res_ant=float(ad.get('resultado_anterior',0) or 0)
            costo_r=spend/res if res>0 else 0
            status=get_status(objetivo,costo_r)
            pct_cambio=((res-res_ant)/res_ant*100) if res_ant>0 else None

            video_plays=get_action(ad.get('video_play_actions',[]),'video_view')
            hook_rate=video_plays/impr*100 if impr>0 and video_plays>0 else 0
            p50=float(ad.get('video_p50',0) or 0)
            p75=float(ad.get('video_p75',0) or 0)
            p100=float(ad.get('video_p100',0) or 0)
            hold50=p50/impr*100 if impr>0 and p50>0 else 0
            hold75=p75/impr*100 if impr>0 and p75>0 else 0
            hold100=p100/impr*100 if impr>0 and p100>0 else 0
            roas_list=ad.get('purchase_roas',[])
            roas=float(roas_list[0].get('value',0)) if roas_list else 0
            clicks=get_action(ad.get('outbound_clicks',[]),'outbound_click')

            bg_row='FFFFFF'
            if status=='⛔ PAUSAR': bg_row='FDEDEC'
            elif status=='🏆 GANADOR': bg_row='EAFAF1'
            elif status=='🔍 REVISAR': bg_row='FEF9E7'

            w(ws_m,mrow,2,ad.get('nombre',''),fg='1E2A38',bg=bg_row)
            w(ws_m,mrow,3,int(reach) if reach else '',bg=bg_row,align='right')
            w(ws_m,mrow,4,int(impr) if impr else '',bg=bg_row,align='right')
            w(ws_m,mrow,5,round(freq,1) if freq else '',bg=bg_row,align='right')
            w(ws_m,mrow,6,int(res) if res else '',bold=True,bg=bg_row,align='right')
            w(ws_m,mrow,7,int(res_ant) if res_ant else '—',italic=bool(res_ant),fg='888888' if not res_ant else '1E2A38',bg=bg_row,align='right')
            if pct_cambio is not None:
                pct_fg='27AE60' if pct_cambio>0 else 'E74C3C'
                w(ws_m,mrow,8,('+' if pct_cambio>0 else '')+f'{pct_cambio:.0f}%',bold=True,fg=pct_fg,bg=bg_row,align='center')
            else:
                w(ws_m,mrow,8,'—',fg='AAAAAA',bg=bg_row,align='center')
            w(ws_m,mrow,9,int(spend) if spend else '',bg=bg_row,align='right',fmt='#,##0')
            w(ws_m,mrow,10,int(costo_r) if costo_r else '',bg=bg_row,align='right',fmt='#,##0')
            w(ws_m,mrow,11,round(ctr,2) if ctr else '',bg=bg_row,align='right')
            w(ws_m,mrow,12,int(cpm) if cpm else '',bg=bg_row,align='right')
            hr_fg='27AE60' if hook_rate>=25 else ('E74C3C' if 0<hook_rate<15 else '1E2A38')
            w(ws_m,mrow,13,round(hook_rate,1) if hook_rate else '',bold=hook_rate>=25,fg=hr_fg,bg=bg_row,align='right')
            w(ws_m,mrow,14,round(hold50,1) if hold50 else '',bg=bg_row,align='right')
            w(ws_m,mrow,15,round(hold75,1) if hold75 else '',bg=bg_row,align='right')
            w(ws_m,mrow,16,round(hold100,1) if hold100 else '',bg=bg_row,align='right')
            w(ws_m,mrow,17,round(roas,2) if roas else '',bg=bg_row,align='right')
            w(ws_m,mrow,18,int(clicks) if clicks else '',bg=bg_row,align='right')
            w(ws_m,mrow,19,status,bold=True,bg=bg_row,align='center',
              fg='27AE60' if 'OK' in status or 'GANADOR' in status else ('E74C3C' if 'PAUSAR' in status else 'E67E22'))
            ws_m.row_dimensions[mrow].height=18
            st_res+=res; st_ant+=res_ant; st_gasto+=spend
            mrow+=1

        ws_m.merge_cells(f'B{mrow}:E{mrow}')
        w(ws_m,mrow,2,f'SUBTOTAL  {camp_name}',bold=True,fg='FFFFFF',bg='2E4057',align='left')
        w(ws_m,mrow,6,int(st_res),bold=True,fg='FFFFFF',bg='2E4057',align='right')
        if st_ant>0:
            w(ws_m,mrow,7,int(st_ant),bold=True,fg='FFFFFF',bg='2E4057',align='right')
            p=((st_res-st_ant)/st_ant*100)
            w(ws_m,mrow,8,('+' if p>0 else '')+f'{p:.0f}%',bold=True,fg='FFFFFF',bg='2E4057',align='center')
        w(ws_m,mrow,9,int(st_gasto),bold=True,fg='FFFFFF',bg='2E4057',align='right',fmt='#,##0')
        ws_m.row_dimensions[mrow].height=20
        mrow+=1
        mrow+=1

    ws_m.freeze_panes='B4'
    ws_m.sheet_view.zoomScale=90

# ── OUTPUT BASE64 ─────────────────────────────────────────────────────────────
buf=io.BytesIO()
wb.save(buf)
buf.seek(0)
print(base64.b64encode(buf.read()).decode('utf-8'))
