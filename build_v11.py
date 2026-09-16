"""
Rebuild v10 → v11: uniform 5 finish surfaces per category.
ЧЕРН stays 3 (Пол/Стены/Потолок).
ЧИСТ becomes 5 everywhere (Пол/Стены/Потолок/Запол./Мебл.).
СПРАВОЧНИК gets Чист:Мебл. column (new c16), ∑ shifts to c17.
SIGNAL formulas updated for new column positions.
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter as CL
from copy import copy
import os

SRC = os.path.join(os.path.dirname(__file__),
                   '02_Финальная версия - Отделка для SIGNAL v10.xlsx')
DST = os.path.join(os.path.dirname(__file__),
                   '03_Финальная версия - Отделка для SIGNAL v11.xlsx')

NC = 3   # черн surfaces
NI = 5   # чист surfaces
CHIST_SURFS = ['Пол', 'Стены', 'Потолок', 'Запол.', 'Мебл.']
CHERN_SURFS = ['Пол', 'Стены', 'Потолок']

# ── styles ──
thin = Side(style='thin')
brd = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
hdr_fill = PatternFill('solid', fgColor='4472C4')
title_font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
title_fill = PatternFill('solid', fgColor='4472C4')
bold_font = Font(name='Calibri', size=11, bold=True)
ug_fill = PatternFill('solid', fgColor='D9E2F3')
ag_fill = PatternFill('solid', fgColor='E2EFDA')
sum_fill = PatternFill('solid', fgColor='FFF2CC')
date_fmt = 'DD.MM.YYYY'
pct_fmt = '0.0%'

def hdr_cell(ws, r, c, val):
    cell = ws.cell(r, c, val)
    cell.font = hdr_font; cell.fill = hdr_fill; cell.border = brd
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    return cell

def data_cell(ws, r, c, val, fill=None, fmt=None, bold=False):
    cell = ws.cell(r, c, val)
    cell.border = brd
    cell.alignment = Alignment(horizontal='center', vertical='center')
    if fill: cell.fill = fill
    if fmt: cell.number_format = fmt
    if bold: cell.font = bold_font
    return cell

def title_cell(ws, r, c, val):
    cell = ws.cell(r, c, val)
    cell.font = title_font; cell.fill = title_fill
    cell.alignment = Alignment(horizontal='left', vertical='center')
    return cell


def read_v10():
    """Extract category data and existing plan/fact values from v10."""
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb['СПРАВОЧНИК']

    # v10 ЧИСТ column widths per category (for reading existing data)
    podzem_chist_widths = [4, 4, 3, 4, 4, 3]  # 8.1.1-8.1.6
    nadzem_chist_widths = [4, 4, 3, 4, 4, 3, 3]  # 8.2.1-8.2.7

    cats = []
    for r in range(3, 16):
        code = ws.cell(r, 1).value
        if not code: continue
        cats.append({
            'code': str(code),
            'part': ws.cell(r, 2).value or '',
            'name': ws.cell(r, 3).value or '',
            'plan_m2': ws.cell(r, 4).value or 0,
            'wc': ws.cell(r, 5).value or 0,
            'wh': ws.cell(r, 6).value or 0,
            'cw': [ws.cell(r, 8).value or 0, ws.cell(r, 9).value or 0, ws.cell(r, 10).value or 0],
            'hw_old': [ws.cell(r, c).value or 0 for c in range(12, 16)],  # Пол,Стены,Пот,Запол
            'start': ws.cell(r, 17).value,
            'end': ws.cell(r, 18).value,
            'dur': ws.cell(r, 19).value or 0,
            'trud': ws.cell(r, 20).value or 0,
            'w3auto': ws.cell(r, 21).value or 0,
            'w3': ws.cell(r, 22).value or 0,
            'gw': ws.cell(r, 23).value or 0,
        })

    # Expand hw_old (4 values) to hw (5 values) — insert Мебл.=0 at end
    # Special case: 8.2.1 (лобби) had [Пол,Стены,Пот,Мебл] → need [Пол,Стены,Пот,0,Мебл]
    for c in cats:
        old = c['hw_old']
        if c['code'] == '8.2.1':
            # v10: Пол, Стены, Потолок, Мебл. → new: Пол, Стены, Потолок, Запол.(0), Мебл.
            c['hw'] = [old[0], old[1], old[2], 0, old[3]]
        else:
            # v10: Пол, Стены, Потолок, Запол. → new: Пол, Стены, Потолок, Запол., Мебл.(0)
            c['hw'] = [old[0], old[1], old[2], old[3], 0]

    podzem = [c for c in cats if c['code'].startswith('8.1.')]
    nadzem = [c for c in cats if c['code'].startswith('8.2.')]

    # Read settings
    settings = {
        'start': ws.cell(20, 2).value,  # approximate rows
        'end': ws.cell(21, 2).value,
        'weeks': 48,
    }
    # Find korpus names and areas
    korpus = []
    areas = {}
    for r in range(16, 36):
        v = ws.cell(r, 1).value
        if v and str(v).startswith('К') and len(str(v)) <= 3:
            k = str(v)
            korpus.append(k)
            a = []
            for c2 in range(2, 2 + len(nadzem)):
                a.append(ws.cell(r, c2).value or 0)
            areas[k] = a

    if not korpus:
        korpus = ['К1', 'К2', 'К3']
        areas = {k: [0]*len(nadzem) for k in korpus}

    # Read plan/fact data from existing sheets
    pfdata = {}
    for sn in wb.sheetnames:
        if not (sn.startswith('ПЛАН') or sn.startswith('ФАКТ')):
            continue
        ws2 = wb[sn]
        is_chist = 'ЧИСТ' in sn
        is_podzem = 'ПОДЗЕМ' in sn
        nc_ = len(podzem) if is_podzem else len(nadzem)
        widths = (podzem_chist_widths if is_podzem else nadzem_chist_widths) if is_chist else [3]*nc_

        rows_data = []
        for r in range(5, ws2.max_row + 1):
            if ws2.cell(r, 1).value is None:
                break
            row = []
            col = 2
            for ci in range(nc_):
                w = widths[ci]
                surfs = []
                for si in range(w):
                    surfs.append(ws2.cell(r, col + si).value or 0)
                # Pad to 5 for chist or 3 for chern
                if is_chist:
                    if len(surfs) == 3:
                        surfs = surfs + [0, 0]  # add Запол., Мебл.
                    elif len(surfs) == 4:
                        if is_podzem or ci != 0:
                            surfs = surfs + [0]  # add Мебл.
                        else:
                            # 8.2.1: was [Пол,Стены,Пот,Мебл] → [Пол,Стены,Пот,0,Мебл]
                            surfs = [surfs[0], surfs[1], surfs[2], 0, surfs[3]]
                row.append(surfs)
                col += w
            rows_data.append(row)
        pfdata[sn] = rows_data

    # Read report date
    sig = wb['SIGNAL']
    report_date = sig.cell(2, 1).value

    wb.close()
    return {
        'podzem': podzem, 'nadzem': nadzem,
        'korpus': korpus, 'areas': areas,
        'pfdata': pfdata, 'report_date': report_date,
        'settings': settings,
    }


def build_v11(data):
    wb = openpyxl.Workbook()
    NP = len(data['podzem'])
    NN = len(data['nadzem'])
    KOR = data['korpus']
    NK = len(KOR)
    W = data['settings']['weeks'] or 48
    all_cats = data['podzem'] + data['nadzem']

    # ── ИНСТРУКЦИЯ ──
    ws = wb.active
    ws.title = 'ИНСТРУКЦИЯ'
    title_cell(ws, 1, 1, 'ИНСТРУКЦИЯ ПО ЗАПОЛНЕНИЮ')
    ws.merge_cells('A1:H1')
    ws.cell(3, 1, 'Файл: Отделка для SIGNAL v11')
    ws.cell(4, 1, 'Черновая отделка: 3 поверхности (Пол, Стены, Потолок)')
    ws.cell(5, 1, 'Чистовая отделка: 5 поверхностей (Пол, Стены, Потолок, Заполнение проёмов, Меблировка)')
    ws.cell(7, 1, 'Справочник — веса и распределения по поверхностям')
    ws.cell(8, 1, 'ПЛАН/ФАКТ листы — ввод м² по неделям')
    ws.cell(9, 1, 'ОТЧЕТ — сводка')
    ws.cell(10, 1, 'SIGNAL — данные для SIGNAL')
    ws.column_dimensions['A'].width = 80

    # ── СПРАВОЧНИК ──
    ws = wb.create_sheet('СПРАВОЧНИК')
    title_cell(ws, 1, 1, 'СПРАВОЧНИК КАТЕГОРИЙ ОТДЕЛКИ')
    ws.merge_cells('A1:X1')

    hdrs = ['Код','Часть','Название','План м²','Вес черн.','Вес чист.','∑ весов',
            'Черн:Пол','Черн:Стены','Черн:Потолок','∑ черн.',
            'Чист:Пол','Чист:Стены','Чист:Потолок','Чист:Запол.','Чист:Мебл.','∑ чист.',
            'Начало','Окончание','Длит.(дни)','Труд-proxy','Вес ур.3 (авто)','Вес ур.3','Глоб. вес']
    for i, h in enumerate(hdrs):
        hdr_cell(ws, 2, i+1, h)

    for idx, c in enumerate(all_cats):
        r = 3 + idx
        fl = ug_fill if c['code'].startswith('8.1.') else ag_fill
        data_cell(ws, r, 1, c['code'], fill=fl)
        data_cell(ws, r, 2, c['part'], fill=fl)
        data_cell(ws, r, 3, c['name'], fill=fl).alignment = Alignment(horizontal='left', wrap_text=True)
        data_cell(ws, r, 4, c['plan_m2'])
        data_cell(ws, r, 5, c['wc'])
        data_cell(ws, r, 6, c['wh'])
        data_cell(ws, r, 7, None).value = f'=E{r}+F{r}'
        for s in range(3):
            data_cell(ws, r, 8+s, c['cw'][s])
        data_cell(ws, r, 11, None).value = f'=SUM(H{r}:J{r})'
        for s in range(5):
            data_cell(ws, r, 12+s, c['hw'][s])
        data_cell(ws, r, 17, None).value = f'=SUM(L{r}:P{r})'
        data_cell(ws, r, 18, c.get('start'))
        data_cell(ws, r, 19, c.get('end'))
        data_cell(ws, r, 20, c.get('dur', 0))
        data_cell(ws, r, 21, c.get('trud', 0))
        data_cell(ws, r, 22, c.get('w3auto', 0))
        data_cell(ws, r, 23, c.get('w3', 0))
        data_cell(ws, r, 24, c.get('gw', 0))

    # Aggregate rows
    agg_r = 3 + len(all_cats) + 1
    data_cell(ws, agg_r, 1, '8.1 Подзем.', bold=True)
    data_cell(ws, agg_r, 4, None).value = f'=SUM(D3:D{2+NP})'
    data_cell(ws, agg_r+1, 1, '8.2 Надзем.', bold=True)
    data_cell(ws, agg_r+1, 4, None).value = f'=SUM(D{3+NP}:D{2+NP+NN})'

    # Korpus block
    kr = agg_r + 3
    data_cell(ws, kr, 1, 'КОРПУСА', bold=True)
    hdr_cell(ws, kr+1, 1, 'Имя')
    for i, nc in enumerate(data['nadzem']):
        hdr_cell(ws, kr+1, 2+i, nc['code'])
    hdr_cell(ws, kr+1, 2+NN, 'Итого')

    for ki, k in enumerate(KOR):
        r_ = kr + 2 + ki
        data_cell(ws, r_, 1, k)
        ar = data['areas'].get(k, [0]*NN)
        for i in range(NN):
            data_cell(ws, r_, 2+i, ar[i] if i < len(ar) else 0)
        data_cell(ws, r_, 2+NN, None).value = f'=SUM(B{r_}:{CL(1+NN)}{r_})'

    # Settings
    sr = kr + 2 + NK + 2
    data_cell(ws, sr, 1, 'Дата начала', bold=True)
    c_sd = data_cell(ws, sr, 2, data['report_date'])
    if c_sd.value: c_sd.number_format = date_fmt
    data_cell(ws, sr+1, 1, 'Дата отчёта', bold=True)
    c_rd = data_cell(ws, sr+1, 2, data['report_date'])
    if c_rd.value: c_rd.number_format = date_fmt
    data_cell(ws, sr+2, 1, 'Недель', bold=True)
    data_cell(ws, sr+2, 2, W)

    # Date row
    dr = sr + 4
    data_cell(ws, dr, 1, 'Даты', bold=True)
    data_cell(ws, dr, 2, data['report_date']).number_format = date_fmt
    for w in range(1, W):
        data_cell(ws, dr, 2+w, None).value = f'={CL(1+w)}{dr}+7'
        ws.cell(dr, 2+w).number_format = date_fmt

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 42
    for i in range(4, 25):
        ws.column_dimensions[CL(i)].width = 11

    sprav_agg_r = agg_r  # remember for SIGNAL formulas

    # ── ПЛАН/ФАКТ sheets ──
    def make_pf_sheet(name, cats, is_chist, existing_data=None):
        ws2 = wb.create_sheet(name)
        ns = NI if is_chist else NC
        surfs = CHIST_SURFS if is_chist else CHERN_SURFS
        total_data_cols = len(cats) * ns
        total_cols = 1 + total_data_cols + ns  # date + data + summary
        part_label = 'Подземная часть' if cats[0]['code'].startswith('8.1.') else 'Надземная часть'
        sum_label = 'Итого'

        # Row 1: title
        title_cell(ws2, 1, 1, name)
        ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)

        # Row 2: part header
        hdr_cell(ws2, 2, 1, 'Дата')
        ws2.merge_cells(start_row=2, start_column=1, end_row=4, end_column=1)
        hdr_cell(ws2, 2, 2, part_label)
        ws2.merge_cells(start_row=2, start_column=2, end_row=2, end_column=1+total_data_cols)

        # Row 3: category names
        for ci, c in enumerate(cats):
            cc = 2 + ci * ns
            lbl = f"{c['code']} {c['name']}"
            hdr_cell(ws2, 3, cc, lbl)
            if ns > 1:
                ws2.merge_cells(start_row=3, start_column=cc, end_row=3, end_column=cc+ns-1)

        # Summary headers
        sc = 2 + total_data_cols
        hdr_cell(ws2, 3, sc, sum_label)
        if ns > 1:
            ws2.merge_cells(start_row=3, start_column=sc, end_row=3, end_column=sc+ns-1)

        # Row 4: surface headers
        for ci in range(len(cats)):
            for si in range(ns):
                hdr_cell(ws2, 4, 2 + ci*ns + si, surfs[si])
        for si in range(ns):
            hdr_cell(ws2, 4, sc + si, surfs[si])

        # Data rows
        for w in range(W):
            r = 5 + w
            if w == 0:
                data_cell(ws2, r, 1, data['report_date']).number_format = date_fmt
            else:
                data_cell(ws2, r, 1, None).value = f'=A{r-1}+7'
                ws2.cell(r, 1).number_format = date_fmt

            for ci in range(len(cats)):
                for si in range(ns):
                    val = 0
                    if existing_data and w < len(existing_data) and ci < len(existing_data[w]):
                        if si < len(existing_data[w][ci]):
                            val = existing_data[w][ci][si]
                    data_cell(ws2, r, 2 + ci*ns + si, val if val else None)

            # Summary formulas
            for si in range(ns):
                parts = [f'{CL(2 + ci*ns + si)}{r}' for ci in range(len(cats))]
                data_cell(ws2, r, sc + si, None, fill=sum_fill).value = '=' + '+'.join(parts)

        ws2.column_dimensions['A'].width = 14
        for i in range(2, total_cols + 1):
            ws2.column_dimensions[CL(i)].width = 10

        return ws2

    sheet_order = []

    # ПЛАН sheets
    make_pf_sheet('ПЛАН ЧЕРН ПОДЗЕМ', data['podzem'], False, data['pfdata'].get('ПЛАН ЧЕРН ПОДЗЕМ'))
    sheet_order.append('ПЛАН ЧЕРН ПОДЗЕМ')
    make_pf_sheet('ПЛАН ЧИСТ ПОДЗЕМ', data['podzem'], True, data['pfdata'].get('ПЛАН ЧИСТ ПОДЗЕМ'))
    sheet_order.append('ПЛАН ЧИСТ ПОДЗЕМ')
    for k in KOR:
        make_pf_sheet(f'ПЛАН ЧЕРН {k}', data['nadzem'], False, data['pfdata'].get(f'ПЛАН ЧЕРН {k}'))
        sheet_order.append(f'ПЛАН ЧЕРН {k}')
        make_pf_sheet(f'ПЛАН ЧИСТ {k}', data['nadzem'], True, data['pfdata'].get(f'ПЛАН ЧИСТ {k}'))
        sheet_order.append(f'ПЛАН ЧИСТ {k}')

    # ФАКТ sheets
    make_pf_sheet('ФАКТ ЧЕРН ПОДЗЕМ', data['podzem'], False, data['pfdata'].get('ФАКТ ЧЕРН ПОДЗЕМ'))
    sheet_order.append('ФАКТ ЧЕРН ПОДЗЕМ')
    make_pf_sheet('ФАКТ ЧИСТ ПОДЗЕМ', data['podzem'], True, data['pfdata'].get('ФАКТ ЧИСТ ПОДЗЕМ'))
    sheet_order.append('ФАКТ ЧИСТ ПОДЗЕМ')
    for k in KOR:
        make_pf_sheet(f'ФАКТ ЧЕРН {k}', data['nadzem'], False, data['pfdata'].get(f'ФАКТ ЧЕРН {k}'))
        sheet_order.append(f'ФАКТ ЧЕРН {k}')
        make_pf_sheet(f'ФАКТ ЧИСТ {k}', data['nadzem'], True, data['pfdata'].get(f'ФАКТ ЧИСТ {k}'))
        sheet_order.append(f'ФАКТ ЧИСТ {k}')

    # ── ОТЧЕТ ──
    ws_rep = wb.create_sheet('ОТЧЕТ')
    title_cell(ws_rep, 1, 1, 'ОТЧЕТ — Сводка')
    ws_rep.merge_cells('A1:F1')
    rep_hdrs = ['Дата', 'Подзем. черн.', 'Подзем. чист.', 'Надзем. черн.', 'Надзем. чист.', 'Итого']
    for i, h in enumerate(rep_hdrs):
        hdr_cell(ws_rep, 3, i+1, h)

    for w in range(W):
        r = 4 + w
        if w == 0:
            data_cell(ws_rep, r, 1, data['report_date']).number_format = date_fmt
        else:
            data_cell(ws_rep, r, 1, None).value = f'=A{r-1}+7'
            ws_rep.cell(r, 1).number_format = date_fmt
        fr = 5 + w
        pc_end = CL(1 + NP * NC)
        pi_end = CL(1 + NP * NI)
        data_cell(ws_rep, r, 2, None).value = f"=SUM('ФАКТ ЧЕРН ПОДЗЕМ'!B{fr}:'ФАКТ ЧЕРН ПОДЗЕМ'!{pc_end}{fr})"
        data_cell(ws_rep, r, 3, None).value = f"=SUM('ФАКТ ЧИСТ ПОДЗЕМ'!B{fr}:'ФАКТ ЧИСТ ПОДЗЕМ'!{pi_end}{fr})"
        ac, ai = [], []
        for k in KOR:
            kc_end = CL(1 + NN * NC)
            ki_end = CL(1 + NN * NI)
            ac.append(f"SUM('ФАКТ ЧЕРН {k}'!B{fr}:'ФАКТ ЧЕРН {k}'!{kc_end}{fr})")
            ai.append(f"SUM('ФАКТ ЧИСТ {k}'!B{fr}:'ФАКТ ЧИСТ {k}'!{ki_end}{fr})")
        data_cell(ws_rep, r, 4, None).value = '=' + ('+'.join(ac) if ac else '0')
        data_cell(ws_rep, r, 5, None).value = '=' + ('+'.join(ai) if ai else '0')
        data_cell(ws_rep, r, 6, None, fill=sum_fill).value = f'=B{r}+C{r}+D{r}+E{r}'

    ws_rep.column_dimensions['A'].width = 14
    for i in range(2, 7):
        ws_rep.column_dimensions[CL(i)].width = 16

    # ── SIGNAL ──
    ws_sig = wb.create_sheet('SIGNAL')
    named_ranges = {}

    def sig_col(block_idx):
        """Each block is 6 cols: label, date, plan, fact, (empty x2). Start at col 2."""
        return 2 + block_idx * 6

    def write_signal_header(bc, title_text):
        """Write standard SIGNAL block header at base column bc."""
        ws_sig.cell(1, bc, 'Заголовок')
        ws_sig.cell(1, bc+1, 'План-факт по объемам')
        ws_sig.cell(2, bc, 'Статус')
        ws_sig.cell(2, bc+1, 'true')
        ws_sig.cell(3, bc, 'Url изображения')
        ws_sig.cell(4, bc, 'Тип')
        ws_sig.cell(4, bc+1, 'planFact2')
        ws_sig.cell(6, bc, 'Дата')
        ws_sig.cell(6, bc+1, '=$A$2')
        ws_sig.cell(7, bc, 'Гистограмма')
        ws_sig.cell(7, bc+1, 'false')
        ws_sig.cell(8, bc, 'По месяцам')
        ws_sig.cell(8, bc+1, 'false')
        ws_sig.cell(10, bc, 'Тип')
        ws_sig.cell(10, bc+1, title_text)
        ws_sig.cell(11, bc, 'Ед. изм.')
        ws_sig.cell(11, bc+1, 'м2')
        ws_sig.cell(12, bc, 'Учитывать дату карточки')
        ws_sig.cell(12, bc+1, 'true')
        ws_sig.cell(13, bc, 'Данные')

    ws_sig.cell(1, 1, 'Дата отчета')
    ws_sig.cell(2, 1, data['report_date'])
    if data['report_date']:
        ws_sig.cell(2, 1).number_format = date_fmt

    # Date row ref
    dr = sr + 4  # from СПРАВОЧНИК

    blocks = []  # list of {bc, name_key, title}

    def add_block(title, name_key):
        idx = len(blocks)
        bc = sig_col(idx)
        blocks.append({'bc': bc, 'name_key': name_key, 'title': title})
        write_signal_header(bc, title)
        return bc

    def weighted_formulas(bc, plan_sheet, fact_sheet, cats, is_chist, sprav_rows_start):
        """Write weighted plan/fact formulas for a block."""
        ns = NI if is_chist else NC
        w_cols = [CL(12+s) for s in range(NI)] if is_chist else [CL(8+s) for s in range(NC)]
        nc_ = len(cats)

        # Row 9: "Всего" — sum of areas
        ws_sig.cell(9, bc, 'Всего')
        area_refs = [f"СПРАВОЧНИК!{CL(2+i)}{sprav_rows_start}" for i in range(nc_)] if not is_chist else []
        # Just use СПРАВОЧНИК Plan m2 sum
        sr_start = 3 + (0 if cats[0]['code'].startswith('8.1.') else NP)
        ws_sig.cell(9, bc+1, f'=SUM(СПРАВОЧНИК!D{sr_start}:D{sr_start+nc_-1})')

        for w in range(W):
            r = 14 + w
            # Date
            if w == 0:
                ws_sig.cell(r, bc).value = '=$A$2'
            else:
                ws_sig.cell(r, bc).value = f'={CL(bc)}{r-1}+7'
            ws_sig.cell(r, bc).number_format = date_fmt

            # Plan formula
            plan_parts = []
            fact_parts = []
            for ci in range(nc_):
                sr_ = sr_start + ci  # СПРАВОЧНИК row for this category
                for si in range(ns):
                    dc = CL(2 + ci * ns + si)
                    wc = w_cols[si]
                    plan_parts.append(f"'{plan_sheet}'!{dc}{5+w}*СПРАВОЧНИК!${wc}${sr_}")
                    fact_parts.append(f"('{fact_sheet}'!{dc}{5+w})*СПРАВОЧНИК!${wc}${sr_}")

            ws_sig.cell(r, bc+1).value = '=' + '+'.join(plan_parts)
            ws_sig.cell(r, bc+2).value = '=' + '+'.join(fact_parts)

    def combined_formulas(bc, chern_bc, chist_bc):
        """Combined = chern + chist."""
        ws_sig.cell(9, bc, 'Всего')
        ws_sig.cell(9, bc+1, f'={CL(chern_bc+1)}9+{CL(chist_bc+1)}9')
        for w in range(W):
            r = 14 + w
            if w == 0:
                ws_sig.cell(r, bc).value = '=$A$2'
            else:
                ws_sig.cell(r, bc).value = f'={CL(bc)}{r-1}+7'
            ws_sig.cell(r, bc).number_format = date_fmt
            ws_sig.cell(r, bc+1).value = f'={CL(chern_bc+1)}{r}+{CL(chist_bc+1)}{r}'
            ws_sig.cell(r, bc+2).value = f'={CL(chern_bc+2)}{r}+{CL(chist_bc+2)}{r}'

    def sum_blocks_formulas(bc, source_bcs):
        """Sum multiple blocks."""
        ws_sig.cell(9, bc, 'Всего')
        ws_sig.cell(9, bc+1, '=' + '+'.join(f'{CL(s+1)}9' for s in source_bcs))
        for w in range(W):
            r = 14 + w
            if w == 0:
                ws_sig.cell(r, bc).value = '=$A$2'
            else:
                ws_sig.cell(r, bc).value = f'={CL(bc)}{r-1}+7'
            ws_sig.cell(r, bc).number_format = date_fmt
            ws_sig.cell(r, bc+1).value = '=' + '+'.join(f'{CL(s+1)}{r}' for s in source_bcs)
            ws_sig.cell(r, bc+2).value = '=' + '+'.join(f'{CL(s+2)}{r}' for s in source_bcs)

    # Block order: same as v10
    # 1) 8.1 Общий (combined chern+chist)
    # 2) 8.1 Черновая (weighted)
    # 3) 8.1 Чистовая (weighted)
    bc_p_o = add_block('8.1 Подземная часть: Общий', 'Общий_8_1')
    bc_p_c = add_block('8.1 Подземная часть: Черновая', 'Черновая_8_1')
    bc_p_i = add_block('8.1 Подземная часть: Чистовая', 'Чистовая_8_1')

    weighted_formulas(bc_p_c, 'ПЛАН ЧЕРН ПОДЗЕМ', 'ФАКТ ЧЕРН ПОДЗЕМ', data['podzem'], False, 3)
    weighted_formulas(bc_p_i, 'ПЛАН ЧИСТ ПОДЗЕМ', 'ФАКТ ЧИСТ ПОДЗЕМ', data['podzem'], True, 3)
    combined_formulas(bc_p_o, bc_p_c, bc_p_i)

    kor_o_bcs, kor_c_bcs, kor_i_bcs = [], [], []
    for k in KOR:
        bc_ko = add_block(f'8.2 Надземная {k}: Общий', f'{k}_Общий_8_2')
        bc_kc = add_block(f'8.2 Надземная {k}: Черновая', f'{k}_Черновая_8_2')
        bc_ki = add_block(f'8.2 Надземная {k}: Чистовая', f'{k}_Чистовая_8_2')

        weighted_formulas(bc_kc, f'ПЛАН ЧЕРН {k}', f'ФАКТ ЧЕРН {k}', data['nadzem'], False, 3+NP)
        weighted_formulas(bc_ki, f'ПЛАН ЧИСТ {k}', f'ФАКТ ЧИСТ {k}', data['nadzem'], True, 3+NP)
        combined_formulas(bc_ko, bc_kc, bc_ki)

        kor_o_bcs.append(bc_ko)
        kor_c_bcs.append(bc_kc)
        kor_i_bcs.append(bc_ki)

    # Aggregate nadzemная
    bc_an_o = add_block('8.2 Надзем. (все): Общий', 'Общий_8_2')
    bc_an_c = add_block('8.2 Надзем. (все): Черновая', 'Черновая_8_2')
    bc_an_i = add_block('8.2 Надзем. (все): Чистовая', 'Чистовая_8_2')
    sum_blocks_formulas(bc_an_o, kor_o_bcs)
    sum_blocks_formulas(bc_an_c, kor_c_bcs)
    sum_blocks_formulas(bc_an_i, kor_i_bcs)

    # Total
    bc_t_o = add_block('8. Отделка общий: Общий', 'Общая')
    bc_t_c = add_block('8. Отделка общий: Черновая', 'Черновая')
    bc_t_i = add_block('8. Отделка общий: Чистовая', 'Чистовая')

    def sum_two_formulas(bc, a_bc, b_bc):
        ws_sig.cell(9, bc, 'Всего')
        ws_sig.cell(9, bc+1, f'={CL(a_bc+1)}9+{CL(b_bc+1)}9')
        for w in range(W):
            r = 14 + w
            if w == 0:
                ws_sig.cell(r, bc).value = '=$A$2'
            else:
                ws_sig.cell(r, bc).value = f'={CL(bc)}{r-1}+7'
            ws_sig.cell(r, bc).number_format = date_fmt
            ws_sig.cell(r, bc+1).value = f'={CL(a_bc+1)}{r}+{CL(b_bc+1)}{r}'
            ws_sig.cell(r, bc+2).value = f'={CL(a_bc+2)}{r}+{CL(b_bc+2)}{r}'

    sum_two_formulas(bc_t_o, bc_p_o, bc_an_o)
    sum_two_formulas(bc_t_c, bc_p_c, bc_an_c)
    sum_two_formulas(bc_t_i, bc_p_i, bc_an_i)

    # Named ranges
    from openpyxl.workbook.defined_name import DefinedName
    for b in blocks:
        bc = b['bc']
        c1 = CL(bc)
        c2 = CL(bc + 3)
        name = b['name_key']
        dn = DefinedName(name, attr_text=f"'SIGNAL'!${c1}:${c2}")
        wb.defined_names.add(dn)

    # Save
    wb.save(DST)
    print(f'Saved: {DST}')
    print(f'Sheets: {len(wb.sheetnames)}')
    print(f'Named ranges: {len(blocks)}')
    print(f'SIGNAL blocks: {len(blocks)}')


if __name__ == '__main__':
    print('Reading v10...')
    d = read_v10()
    print(f'Categories: {len(d["podzem"])} podzem + {len(d["nadzem"])} nadzem')
    print(f'Korpus: {d["korpus"]}')
    print('Building v11...')
    build_v11(d)
    print('Done!')
