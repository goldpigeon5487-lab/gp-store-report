import io
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='GP 櫃位報表轉換', layout='wide')
st.title('GP 櫃位訂單報表轉換')
st.caption('依廠商編號彙整尺寸；取消訂單排除；缺漏資料提供檢查清單。')

ALIASES = {
    'store': ['POS商店名稱', 'POS 商店名稱', '櫃位名稱', '商店名稱'],
    'vendor': ['廠商編號', '廠商商品編號', '廠商料號'],
    'qty': ['數量', '商品數量', '購買數量'],
    'order': ['訂單編號', '訂單號碼', '訂單ID', 'Order ID', 'Order Number'],
    'status': ['訂單狀態', '訂單狀態名稱', 'Order Status'],
    'cancel_time': ['取消時間', '取消日期', '訂單取消時間', 'Cancelled At', 'Canceled At'],
    'product': ['商品款式', '商品名稱', '品名'],
}
CANCEL_VALUES = {'已取消', '取消', '訂單已取消', '已作廢', '作廢', 'cancelled', 'canceled'}
APPAREL = ['XS', 'S', 'M', 'L', 'XL', 'XXL']


def norm(v):
    if pd.isna(v):
        return ''
    return re.sub(r'\s+', '', str(v).replace('\ufeff', '').replace('\u3000', ' ')).lower()


def clean(v):
    return '' if pd.isna(v) else str(v).strip()


def detect_header(raw):
    best_row, best_score = 0, -1
    for i in range(min(30, len(raw))):
        vals = {norm(x) for x in raw.iloc[i].tolist()}
        score = sum((3 if k in ('store', 'vendor', 'qty') else 1)
                    for k, al in ALIASES.items() if vals.intersection({norm(a) for a in al}))
        if score > best_score:
            best_row, best_score = i, score
    return best_row


def load_raw(upload, sheet):
    content = io.BytesIO(upload.getvalue())
    if upload.name.lower().endswith('.csv'):
        for encoding in ('utf-8-sig', 'cp950', 'big5'):
            try:
                content.seek(0)
                return pd.read_csv(content, header=None, dtype=str, keep_default_na=False, encoding=encoding)
            except UnicodeDecodeError:
                continue
        raise ValueError('CSV 編碼無法辨識，請先另存為 UTF-8 或 XLSX。')
    return pd.read_excel(content, sheet_name=sheet, header=None, dtype=str, keep_default_na=False)


def prepare(raw, header_row):
    df = raw.iloc[header_row + 1:].copy().reset_index(drop=True)
    names = [clean(x) or f'未命名欄位_{i+1}' for i, x in enumerate(raw.iloc[header_row])]
    # 保證欄名不重複；後續用下拉選欄位才不會含糊。
    seen = defaultdict(int)
    unique = []
    for name in names:
        seen[name] += 1
        unique.append(name if seen[name] == 1 else f'{name} ({seen[name]})')
    df.columns = unique
    df = df.loc[~df.apply(lambda r: all(clean(x) == '' for x in r), axis=1)].copy()
    df['_原始Excel列號'] = df.index + header_row + 2
    return df


def suggest_column(columns, key):
    lookup = {norm(c): c for c in columns}
    for alias in ALIASES[key]:
        if norm(alias) in lookup:
            return lookup[norm(alias)]
    return None


def safe_sheet_name(name, existing):
    base = re.sub(r'[\\/*?:\[\]]', '_', str(name)).strip(" '\t") or '未命名櫃位'
    base = base[:31]
    candidate, n = base, 2
    while candidate.lower() in {x.lower() for x in existing}:
        suffix = f'_{n}'
        candidate = base[:31-len(suffix)] + suffix
        n += 1
    existing.add(candidate)
    return candidate


def parse_size(code):
    # 只有最後一個 _ 後面的部分是尺寸；商品名稱 / 款式不參與解析。
    if '_' not in code:
        return code, None, None
    model, size = code.rsplit('_', 1)
    size = size.strip().upper()
    if not model.strip() or not size:
        return model, None, '廠商編號尺寸格式不完整'
    if size in APPAREL:
        return model, 3 + APPAREL.index(size), None  # XS 對應第 4 格 (Excel E)
    model_part = model.rsplit('-', 1)[0] if '-' in model else model
    m = re.search(r'([A-Za-z]+)$', model_part)
    suffix = m.group(1).upper() if m else ''
    if not size.isdigit():
        return model, None, f'無法辨識的尺寸：{size}'
    value = int(size)
    if suffix == 'B' and 24 <= value <= 40:
        return model, value - 24, None
    if (suffix in ('M', 'W', 'BW', '') or suffix not in ('B',)) and 34 <= value <= 48:
        return model, value - 34, None
    return model, None, f'尺寸不在既有範圍：{size}'


def calculate(df, col, cancel_states):
    data = df.copy()
    order = data[col['order']].map(clean)
    status = data[col['status']].map(clean)
    cancel_norm = {norm(x) for x in cancel_states}
    direct_cancel = status.map(norm).isin(cancel_norm)
    if col.get('cancel_time'):
        direct_cancel |= data[col['cancel_time']].map(clean).ne('')
    # 同編號所有商品皆排除；訂單編號空白時只排除當筆。
    ids = set(order[direct_cancel & order.ne('')])
    cancelled = direct_cancel | (order.ne('') & order.isin(ids))
    dropped = data.loc[cancelled].copy()
    valid = data.loc[~cancelled].copy()
    warnings = []
    specials = []
    accum = defaultdict(lambda: defaultdict(lambda: {'sizes': [0] * 17, 'total': 0, 'no_size': 0, 'unmapped': 0}))

    for _, r in valid.iterrows():
        row_no = int(r['_原始Excel列號'])
        store = clean(r[col['store']])
        code = clean(r[col['vendor']])
        order_id = clean(r[col['order']])
        qty_raw = clean(r[col['qty']])
        product = clean(r[col['product']]) if col.get('product') else ''

        def issue(reason):
            warnings.append({'原始Excel列號': row_no, '訂單編號': order_id, '櫃位': store,
                             '廠商編號': code, '商品款式': product, '數量原值': qty_raw, '問題': reason})

        if not store:
            issue('POS商店名稱空白（未列入彙整）')
            continue
        if not code:
            issue('廠商編號空白（未列入彙整，請回原系統補正）')
            continue
        value = pd.to_numeric(qty_raw.replace(',', ''), errors='coerce')
        if pd.isna(value) or not float(value).is_integer():
            issue('數量不是有效整數（未列入彙整）')
            continue
        qty = int(value)
        model, index, error = parse_size(code)
        if error:
            if re.fullmatch(r'\d+\s*/\s*\d+', code.rsplit('_', 1)[-1]) and '_' in code:
                specials.append({'櫃位': store, '商品型號-色號': model, '特殊尺寸': code.rsplit('_', 1)[-1], '數量': qty, '訂單編號': order_id})
            else:
                issue(error + '（列入合計但沒有尺寸格）')
        bucket = accum[store][model]
        bucket['total'] += qty
        if '_' not in code:
            bucket['no_size'] += qty
        elif index is None:
            bucket['unmapped'] += qty
        else:
            bucket['sizes'][index] += qty

    tables = {}
    for store, models in sorted(accum.items()):
        rows = []
        for model, info in sorted(models.items()):
            sizes = [x if x != 0 else '' for x in info['sizes']]
            rows.append([model] + sizes + [info['total']])
        tables[store] = pd.DataFrame(rows, columns=['商品型號-色號'] + [f'位置{i}' for i in range(17)] + ['合計'])
    special_df = pd.DataFrame(specials, columns=['櫃位', '商品型號-色號', '特殊尺寸', '數量', '訂單編號'])
    return valid, dropped, pd.DataFrame(warnings, columns=['原始Excel列號', '訂單編號', '櫃位', '廠商編號', '商品款式', '數量原值', '問題']), special_df, tables


def style_basic(ws):
    ws.freeze_panes = 'A2'
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor='DDEFE5')
        cell.font = Font(bold=True)
    for j in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(j)].width = 22 if j <= 5 else 16


def export_excel(original, valid, dropped, warnings, specials, tables):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        for name, data in [('訂單明細', original), ('有效訂單明細', valid), ('排除取消訂單', dropped), ('資料異常', warnings), ('特殊尺寸明細', specials)]:
            data.drop(columns=['_原始Excel列號'], errors='ignore').to_excel(writer, sheet_name=name, index=False)
            style_basic(writer.sheets[name])
        used = set(writer.sheets)
        for store, table in tables.items():
            name = safe_sheet_name(store, used)
            # 保留原版 19 欄版型（A 型號，B:R 尺寸，S 合計）
            table.to_excel(writer, sheet_name=name, startrow=3, index=False, header=False)
            ws = writer.sheets[name]
            ws['A1'] = store
            ws.merge_cells('A1:A2')
            for i in range(15):
                ws.cell(1, i + 2, str(34 + i))
            for i in range(17):
                ws.cell(2, i + 2, str(24 + i))
            ws['A3'] = '商品型號-色號'
            for i, s in enumerate(APPAREL):
                ws.cell(3, i + 5, s)
            ws['S3'] = '合計'
            ws.freeze_panes = 'B4'
            ws.sheet_view.showGridLines = False
            ws.column_dimensions['A'].width = 21
            for i in range(2, 20):
                ws.column_dimensions[get_column_letter(i)].width = 6.5
            ws.column_dimensions['S'].width = 9
            edge = Side(style='thin', color='CCD6D2')
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=19):
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    if cell.row <= 3:
                        cell.fill = PatternFill('solid', fgColor='CCFFCC')
                        cell.font = Font(bold=True)
                    if cell.row >= 4 or cell.value is not None:
                        cell.border = Border(left=edge, right=edge, top=edge, bottom=edge)
        # 將各櫃位的轉換結果置於 Excel 最前方，其餘稽核明細排在後面。
        # 只調整工作表順序，不影響任何數值、公式或樣式。
        result_names = [name for name in writer.book.sheetnames if name not in {
            '訂單明細', '有效訂單明細', '排除取消訂單', '資料異常', '特殊尺寸明細'
        }]
        audit_names = ['資料異常', '特殊尺寸明細', '排除取消訂單', '有效訂單明細', '訂單明細']
        ordered_names = result_names + [name for name in audit_names if name in writer.book.sheetnames]
        writer.book._sheets = [writer.book[name] for name in ordered_names]
        writer.book.active = 0
    return out.getvalue()


uploaded = st.file_uploader('上傳 CYBERBIZ 訂單報表（XLSX／CSV）', type=['xlsx', 'csv'])
if uploaded:
    try:
        if uploaded.name.lower().endswith('.xlsx'):
            sheets = pd.ExcelFile(io.BytesIO(uploaded.getvalue())).sheet_names
            sheet = st.selectbox('選擇來源分頁', sheets)
        else:
            sheet = None
        raw = load_raw(uploaded, sheet)
        found = detect_header(raw)
        header_row = st.number_input('表頭所在列（可手動修正）', min_value=1,
                                     max_value=max(1, len(raw)), value=found+1, step=1) - 1
        data = prepare(raw, int(header_row))
        if data.empty:
            st.error('表頭下方沒有明細資料。')
            st.stop()
        options = ['（不使用）'] + [c for c in data.columns if c != '_原始Excel列號']
        st.subheader('1. 確認欄位對應')
        keys = [('store', 'POS商店名稱'), ('vendor', '廠商編號'), ('qty', '數量'),
                ('order', '訂單編號'), ('status', '訂單狀態'), ('cancel_time', '取消時間（選用）'),
                ('product', '商品名稱／款式（選用）')]
        col = {}
        controls = st.columns(3)
        for n, (key, label) in enumerate(keys):
            proposed = suggest_column(data.columns, key)
            with controls[n % 3]:
                selected = st.selectbox(label, options, index=options.index(proposed) if proposed else 0, key=f'map_{key}')
                col[key] = None if selected == '（不使用）' else selected
        missing = [label for key, label in keys[:5] if not col[key]]
        if missing:
            st.error('請指定必要欄位：' + '、'.join(missing))
            st.stop()
        if len([col[k] for k in ('store', 'vendor', 'qty', 'order', 'status')]) != len(set(col[k] for k in ('store', 'vendor', 'qty', 'order', 'status'))):
            st.error('必要欄位不可重複指定。')
            st.stop()
        states = sorted({clean(v) for v in data[col['status']] if clean(v)})
        chosen = st.multiselect('視為「取消」的訂單狀態', states,
                                default=[v for v in states if norm(v) in {norm(x) for x in CANCEL_VALUES}])
        if not chosen and not col['cancel_time']:
            st.warning('未選擇取消狀態，也沒有取消時間欄位；取消訂單可能無法排除。')
        valid, dropped, warnings, specials, tables = calculate(data, col, chosen)
        st.subheader('2. 轉換結果')
        a, b, c, d = st.columns(4)
        a.metric('原始明細筆數', len(data))
        b.metric('取消明細筆數', len(dropped))
        c.metric('有效明細筆數', len(valid))
        d.metric('資料異常筆數', len(warnings))
        if not warnings.empty:
            st.warning(f'有 {len(warnings)} 筆資料異常，請回 CYBERBIZ 補正後重新匯出。')
            with st.expander('查看資料異常', expanded=True):
                st.dataframe(warnings, hide_index=True, use_container_width=True)
        if not specials.empty:
            st.info(f'另有 {len(specials)} 筆複合尺寸商品（例如 43/44），數量已計入合計，詳細尺寸請見「特殊尺寸明細」。')
            with st.expander('查看特殊尺寸明細'):
                st.dataframe(specials, hide_index=True, use_container_width=True)
        if tables:
            choice = st.selectbox('預覽櫃位', list(tables))
            preview = tables[choice].copy()
            preview.columns = ['商品型號-色號'] + [str(i) for i in range(24, 41)] + ['合計']
            st.caption('尺寸欄位置與原 Excel 相同：第 1～15 個位置可對應 34～48；第 1～17 個位置可對應 24～40。服飾 XS～XXL 依既有設定放在第 4～9 個位置。')
            st.dataframe(preview, use_container_width=True, hide_index=True)
        else:
            st.info('沒有可彙整的有效商品，仍可下載明細與資料異常。')
        st.subheader('3. 下載 Excel')
        xlsx = export_excel(data, valid, dropped, warnings, specials, tables)
        filename = datetime.now().strftime('%Y%m%d') + '櫃位轉檔.xlsx'
        st.download_button('下載完整報表 Excel', data=xlsx, file_name=filename,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
    except Exception as e:
        st.error(f'讀取或轉換失敗：{e}')
        st.exception(e)
