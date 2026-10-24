import csv, re, sys, collections, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

csv.field_size_limit(10**8)
D = sys.argv[1]
OUT = sys.argv[2]
TODAY = dt.date(2026, 10, 24)
C = collections.Counter

def nz(v):
    v = (v or '').strip()
    return '' if v in ('NULL', 'None') else v

# ---------- reference ----------
ref = {r['ID']: r['post_title'] for r in csv.DictReader(open(f'{D}/ref.tsv'), delimiter='\t', quoting=csv.QUOTE_NONE)}
LINE_WP = {
    'Từ Việt Nam sang Nhật Bản': ('Việt → Nhật', 'Nhật Bản'), 'Từ Việt Nam sang Đức': ('Việt → Đức', 'Đức'),
    'Từ Đức Về Việt Nam': ('Đức → Việt', 'Việt Nam'), 'Từ Nhật Bản về Việt Nam': ('Nhật → Việt', 'Việt Nam'),
    'Từ Việt Nam sang Châu Âu': ('Việt → Âu', ''), 'Từ Việt Nam sang Séc': ('Việt → Séc', 'Séc'),
    'Từ Việt Nam sang Mỹ': ('Việt → Mỹ', 'Mỹ'), 'Từ Việt Nam sang Canada': ('Việt → Canada', 'Canada'),
    'Từ Việt Nam sang Úc': ('Việt → Úc', 'Úc'), 'Từ Mỹ sang Việt Nam': ('Mỹ → Việt', 'Việt Nam'),
    'Từ Việt Nam sang Pháp': ('Việt → Pháp', 'Pháp'), 'Từ Việt Nam sang Đài Loan': ('Việt → Đài Loan', 'Đài Loan'),
    'Từ Việt Nam sang Hàn Quốc': ('Việt → Hàn', 'Hàn Quốc'), 'Từ Úc Về Việt Nam': ('Úc → Việt', 'Việt Nam'),
}
LINE_ERP = {
    'VIET_DUC': 'Việt → Đức', 'JP_VIET': 'Nhật → Việt', 'DUC_VIET': 'Đức → Việt', 'VIET_UC': 'Việt → Úc',
    'VIET_MY': 'Việt → Mỹ', 'VIET_AU': 'Việt → Âu', 'VIET_PHAP': 'Việt → Pháp', 'VIET_SEC': 'Việt → Séc',
    'VIET_CANADA': 'Việt → Canada', 'UC_VIET': 'Úc → Việt',
}
LINE_DEST = {v[0]: v[1] for v in LINE_WP.values()}
NATION = {
    'germany': 'Đức', 'deutschland': 'Đức', 'đức': 'Đức', 'france': 'Pháp', 'pháp': 'Pháp', 'japan': 'Nhật Bản',
    'nhật': 'Nhật Bản', 'nhật bản': 'Nhật Bản', 'vietnam': 'Việt Nam', 'viet nam': 'Việt Nam', 'việt nam': 'Việt Nam',
    'czech republic': 'Séc', 'czechia': 'Séc', 'séc': 'Séc', 'netherlands': 'Hà Lan', 'poland': 'Ba Lan',
    'sweden': 'Thụy Điển', 'finland': 'Phần Lan', 'slovakia': 'Slovakia', 'hungary': 'Hungary', 'belgium': 'Bỉ',
    'spain': 'Tây Ban Nha', 'denmark': 'Đan Mạch', 'italy': 'Ý', 'united kingdom': 'Anh', 'austria': 'Áo',
    'switzerland': 'Thụy Sĩ', 'luxembourg': 'Luxembourg', 'usa': 'Mỹ', 'united states': 'Mỹ', 'mỹ': 'Mỹ',
    'australia': 'Úc', 'úc': 'Úc', 'canada': 'Canada', 'hàn quốc': 'Hàn Quốc', 'korea': 'Hàn Quốc',
    'đài loan': 'Đài Loan', 'taiwan': 'Đài Loan', 'portugal': 'Bồ Đào Nha', 'norway': 'Na Uy', 'ireland': 'Ireland',
    'greece': 'Hy Lạp', 'romania': 'Romania', 'bulgaria': 'Bulgaria', 'slovenia': 'Slovenia', 'croatia': 'Croatia',
    'estonia': 'Estonia', 'latvia': 'Latvia', 'lithuania': 'Lithuania',
}

def norm_nation(v, line):
    v = nz(v)
    if v:
        return NATION.get(v.lower(), v[:1].upper() + v[1:])
    return LINE_DEST.get(line, '') or '(chưa rõ)'

def norm_code(v):
    s = nz(v).upper().replace(' ', ' ').strip()
    if not s:
        return ''
    s = re.sub(r'^AMAMY[\s\-_]*', '', s)
    m = re.match(r'^([0-9]*[A-Z][A-Z\-]*?[A-Z])\s*-?\s*(\d+)', s)
    if m:
        return m.group(1) + m.group(2)
    return re.sub(r'\s+', '', s)

def norm_phone(v):
    d = re.sub(r'\D', '', nz(v))
    if len(d) < 8 or len(set(d)) <= 2 or d.endswith('123456789'):
        return ''
    return d[-9:]

def norm_email(v):
    e = nz(v).lower()
    return e if re.match(r'^[^@\s]+@[^@\s]+\.[a-z]{2,}$', e) else ''

def clean_city(v):
    s = nz(v)
    s = re.sub(r'\b\d{4,6}(-\d{4})?\b', '', s)  # bỏ mã bưu chính
    s = re.sub(r'\s+', ' ', s).strip(' ,.-')
    if not s or re.fullmatch(r'[\W\d_]*', s):
        return '(chưa rõ)'
    return s[:1].upper() + s[1:]

TEST_RE = re.compile(r'\btest\b|nguyenvana|nguyễn văn a\b', re.I)
COMPANY_EMAIL = re.compile(r'amamy|okhub', re.I)

# ---------- website waybills ----------
status_title = ref
wp = []
n_test_wp = 0
for r in csv.DictReader(open(f'{D}/mvd.tsv'), delimiter='\t', quoting=csv.QUOTE_NONE):
    lt = ref.get(nz(r['chieu_van_don']), '')
    line = LINE_WP.get(lt, ('(chưa rõ)', ''))[0]
    code_raw = nz(r['ma_khach_hang'])
    name = nz(r['ten_nguoi_gui']) or nz(r['name_facebook'])
    if TEST_RE.search(code_raw) or TEST_RE.search(name):
        n_test_wp += 1
        continue
    email = norm_email(r['user'])
    if COMPANY_EMAIL.search(email):
        email = ''
    wp.append(dict(
        src='Website', id=r['ID'], date=r['post_date'][:10], line=line, code_raw=code_raw, code=norm_code(code_raw),
        name=name, phone_raw=nz(r['sdt']) or nz(r['sdt_nguoi_gui']), phone=norm_phone(r['sdt']) or norm_phone(r['sdt_nguoi_gui']),
        email=email, recv=nz(r['ten_nguoi_nhan']), nation=norm_nation(r['nation'], line),
        city=clean_city(r['tinh_thanh_nguoi_nhan']), status=status_title.get(nz(r['trang_thai_don_hang']), ''),
        tracking=nz(r['title']), kg=None, money=None, packages=None))

# ---------- ERP orders (kg / tiền thật) ----------
erp = []
for r in csv.DictReader(open(f'{D}/erp.csv')):
    if r['line'] not in LINE_ERP or r['wrong_marked_at']:
        continue
    if TEST_RE.search(r['customer_code'] or '') or TEST_RE.search(r['sender_name'] or '') or TEST_RE.search(r['goods_note'] or ''):
        continue
    line = LINE_ERP[r['line']]
    kg = float(r['confirmed_weight_kg'] or r['weight_kg'] or 0)
    money = float(r['revenue'] or r['bill_total'] or 0) + float(r['extra_payment_amount'] or 0)
    # giống báo cáo ERP: email người gửi, thiếu thì lấy email người nhận
    email = norm_email(r['sender_email']) or norm_email(r.get('recipient_email'))
    if COMPANY_EMAIL.search(email):
        email = ''
    erp.append(dict(
        src='ERP', id=r['id'][:8], date=r['packing_date'], line=line, code_raw=r['customer_code'] or '',
        code=norm_code(r['customer_code']), name=nz(r['sender_name']) or nz(r['recipient_name']), phone_raw=nz(r['sender_phone']) or nz(r.get('recipient_phone')),
        phone=norm_phone(r['sender_phone']) or norm_phone(r.get('recipient_phone')), email=email, recv=nz(r['recipient_name']),
        nation=norm_nation(r['recipient_nation'], line), city=clean_city(r['recipient_province']),
        status=r['shipping_status'] or '', tracking='', kg=kg, money=money, packages=int(r['packages'] or 0)))

# ---------- khach-hang CPT: kênh, nhân viên ----------
cpt = {}
KENH = {'fanpage': 'Fanpage', 'fb_van_chuyen': 'FB vận chuyển', 'tiktok': 'TikTok', 'titktok': 'TikTok', 'zalo_ca_nhan': 'Zalo cá nhân'}
for r in csv.DictReader(open(f'{D}/kh.tsv'), delimiter='\t', quoting=csv.QUOTE_NONE):
    c = norm_code(nz(r['ma_kh']) or r['post_title'])
    if c and not TEST_RE.search(c):
        cpt[c] = dict(kenh=KENH.get(nz(r['kenh']), nz(r['kenh'])), nv=nz(r['nv']), created=r['post_date'][:10])

# ---------- ID khách = email (trường "user" trên mã vận đơn — website cũng tra cứu khách theo email này) ----------
allo = wp + erp
# đơn thiếu email: mượn email từ đơn khác có cùng mã KH, rồi cùng SĐT (chỉ khi mã/SĐT đó chỉ gắn với 1 email)
def email_map(field):
    m = collections.defaultdict(C)
    for o in allo:
        if o['email'] and o[field]: m[o[field]][o['email']] += 1
    return {k: v.most_common(1)[0][0] for k, v in m.items() if len(v) == 1}
by_code, by_phone = email_map('code'), email_map('phone')
for o in allo:
    o['email_src'] = 'Trên đơn' if o['email'] else ''
    if not o['email'] and o['code'] in by_code:
        o['email'], o['email_src'] = by_code[o['code']], 'Theo mã KH'
    if not o['email'] and o['phone'] in by_phone:
        o['email'], o['email_src'] = by_phone[o['phone']], 'Theo SĐT'
    # không có email: tách riêng theo mã KH / SĐT / tên
    o['_k'] = o['email'] or ('(không email) ' + (o['code'] or o['phone'] or o['name'].lower() or f'#{o["src"]}{o["id"]}'))

groups = collections.defaultdict(list)
for o in allo:
    groups[o['_k']].append(o)

def make_orders(g):
    """Gom bản ghi của 1 khách thành đơn (1 đơn = 1 lần gửi):
    - website: cùng line, cách nhau ≤14 ngày = 1 đơn (form đặt đơn + mã vận đơn tạo sau + nhiều thùng)
    - đơn có cân/tiền (line_orders) được gắn vào đơn website cùng line gần ngày nhất (±21 ngày), không có thì thành đơn riêng."""
    orders = []
    by = collections.defaultdict(list)
    for o in g:
        if o['src'] == 'Website': by[o['line']].append(o)
    for line, recs in by.items():
        recs.sort(key=lambda o: o['date'] or '')
        cur = None
        for o in recs:
            d = dt.date.fromisoformat(o['date']) if o['date'] else None
            if cur is None or d is None or cur['_d0'] is None or (d - cur['_d0']).days > 14:
                cur = dict(line=line, recs=[], paid=[], _d0=d)
                orders.append(cur)
            cur['recs'].append(o)
    for e in sorted((o for o in g if o['src'] == 'ERP'), key=lambda o: o['date']):
        d = dt.date.fromisoformat(e['date'])
        best = None
        for od in orders:
            if od['line'] != e['line'] or od['paid'] or od['_d0'] is None:
                continue
            gap = abs((d - od['_d0']).days)
            if gap <= 21 and (best is None or gap < best[0]):
                best = (gap, od)
        if best:
            best[1]['paid'].append(e)
        else:
            orders.append(dict(line=e['line'], recs=[], paid=[e], _d0=d))
    out = []
    for od in orders:
        allr = od['paid'] + od['recs']
        coded = [o for o in allr if o['code_raw']]
        first = (coded or allr)[0]
        dates = [o['date'] for o in od['paid']] or sorted(o['date'] for o in od['recs'] if o['date'])
        trk = [o['tracking'] for o in od['recs'] if o['tracking'] and not o['tracking'].startswith(o['code_raw'] or '\0')]
        st = [o['status'] for o in od['recs'] if o['status']] + [o['status'] for o in od['paid'] if o['status']]
        has_paid = bool(od['paid'])
        out.append(dict(
            date=dates[0] if dates else '', line=od['line'], code_raw=first['code_raw'], name=first['name'] or next((o['name'] for o in allr if o['name']), ''),
            phone_raw=next((o['phone_raw'] for o in allr if o['phone_raw']), ''), recv=next((o['recv'] for o in allr if o['recv']), ''),
            nation=top([o['nation'] for o in allr]), city=top([o['city'] for o in allr]),
            tracking=', '.join(dict.fromkeys(trk)), n_box=len(od['recs']) or None,
            packages=sum(o['packages'] for o in od['paid']) if has_paid else (len(trk) or None),
            kg=round(sum(o['kg'] for o in od['paid']), 2) if has_paid else None,
            money=sum(o['money'] for o in od['paid']) if has_paid else None,
            status=st[-1] if st else ''))
    return out

def top(vals):
    vals = [v for v in vals if v and v != '(chưa rõ)']
    return C(vals).most_common(1)[0][0] if vals else '(chưa rõ)'

def joined(vals, n=4):
    seen = [v for v, _ in C(v for v in vals if v).most_common()]
    return ', '.join(seen[:n]) + (f' (+{len(seen) - n})' if len(seen) > n else '')

def src_code(code):
    # mã nguồn = nhóm chữ sau F trong mã KH (vd DEFPT193 -> PT), nếu không có thì để trống
    m = re.search(r'F([A-Z]{2})\d', code or '')
    return m.group(1) if m else ''

custs, orders_all = [], []
for g in groups.values():
    ords = make_orders(g)
    codes = [o['code'] for o in g if o['code']]
    main_code = top(codes) if codes else ''
    info = next((cpt[c] for c in [main_code] + codes if c in cpt), {})
    dates = sorted(o['date'] for o in ords if o['date'])
    paid = [o for o in ords if o['money'] is not None]
    c = dict(
        key=g[0]['_k'], has_email=bool(g[0]['email']),
        code=main_code, codes=joined(codes, 3), name=top([o['name'] for o in g]),
        phone=top([o['phone_raw'] for o in g if o['phone']]) if any(o['phone'] for o in g) else '',
        email=g[0]['email'],
        line=top([o['line'] for o in ords]), lines=joined([o['line'] for o in ords]), n_lines=len({o['line'] for o in ords}),
        nation=top([o['nation'] for o in ords]), city=top([o['city'] for o in ords]),
        kenh=info.get('kenh', ''), nv=info.get('nv', '') or src_code(main_code),
        n_orders=len(ords), n_paid=len(paid),
        kg=round(sum(o['kg'] for o in paid), 2), money=sum(o['money'] for o in paid),
        first=dates[0] if dates else '', last=dates[-1] if dates else '')
    c['days'] = (TODAY - dt.date.fromisoformat(c['last'])).days if c['last'] else None
    for o in ords:
        o['_cust'] = c
    custs.append(c)
    orders_all.extend(ords)

# ---------- phân loại ----------
money_vals = sorted([c['money'] for c in custs if c['money'] > 0], reverse=True)
top10 = money_vals[max(0, len(money_vals) // 10 - 1)] if money_vals else 0
for c in custs:
    n, d, m = c['n_orders'], c['days'] or 0, c['money']
    if n >= 5 or (n >= 2 and m >= top10 and m > 0):
        seg = 'VIP'
    elif n >= 2 and d > 90:
        seg = 'Cần chăm sóc lại (≥2 đơn, >90 ngày)'
    elif n >= 2:
        seg = 'Thân thiết'
    elif m >= top10 and m > 0:
        seg = 'Giá trị cao (1 đơn)'
    elif d <= 60:
        seg = 'Mới (1 đơn, ≤60 ngày)'
    else:
        seg = 'Một lần, đã lâu (>60 ngày)'
    c['seg'] = seg

custs.sort(key=lambda c: (-c['money'], -c['n_orders'], -c['kg'], c['last'] or ''))
for i, c in enumerate(custs, 1):
    c['id'] = c['key']
    c['rank'] = i

# ================= WORKBOOK =================
wb = Workbook()
F = 'Arial'
HFONT = Font(name=F, bold=True, color='FFFFFF')
HFILL = PatternFill('solid', fgColor='1F4E78')
BODY = Font(name=F, size=10)
VND = '#,##0;(#,##0);-'
KG = '#,##0.0;(#,##0.0);-'
INT = '#,##0;(#,##0);-'
PCT = '0.0%;(0.0%);-'
tables = 0

def sheet(ws, headers, rows, fmts, widths):
    global tables
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font, c.fill = HFONT, HFILL
        c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[1].height = 32
    for j, (f, w) in enumerate(zip(fmts, widths), 1):
        col = get_column_letter(j)
        ws.column_dimensions[col].width = w
        for (cell,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
            cell.font = BODY
            if f: cell.number_format = f
    ws.freeze_panes = 'A2'
    tables += 1
    t = Table(displayName=f'T{tables}', ref=f'A1:{get_column_letter(len(headers))}{len(rows) + 1}')
    t.tableStyleInfo = TableStyleInfo(name='TableStyleLight9', showRowStripes=True)
    ws.add_table(t)

# --- Khách hàng ---
ws = wb.active
ws.title = 'Khách hàng'
H = ['Hạng', 'ID khách (email)', 'Phân loại', 'Mã KH chính', 'Các mã KH', 'Tên khách', 'SĐT', 'Email',
     'Line chính', 'Các line đã dùng', 'Số line', 'Quốc gia nhận', 'Thành phố nhận', 'Kênh', 'Mã nguồn / NV',
     'Số đơn', 'Số đơn có kg & tiền', 'Tổng kg', 'Tổng tiền (VNĐ)', 'TB / đơn (VNĐ)', '% lũy kế doanh thu',
     'Đơn đầu', 'Đơn gần nhất', 'Số ngày từ đơn gần nhất']
n = len(custs)
rows = []
for i, c in enumerate(custs, 2):
    rows.append([c['rank'], c['id'], c['seg'], c['code'], c['codes'], c['name'], c['phone'], c['email'],
                 c['line'], c['lines'], c['n_lines'], c['nation'], c['city'], c['kenh'], c['nv'],
                 c['n_orders'], c['n_paid'], c['kg'], c['money'],
                 f'=IF(Q{i}>0,S{i}/Q{i},0)', f'=IF(SUM($S$2:$S${n + 1})>0,SUM($S$2:S{i})/SUM($S$2:$S${n + 1}),0)',
                 c['first'], c['last'], c['days']])
sheet(ws, H, rows,
      [INT, None, None, '@', None, None, '@', None, None, None, INT, None, None, None, None,
       INT, INT, KG, VND, VND, PCT, None, None, INT],
      [6, 30, 22, 14, 22, 24, 14, 26, 14, 26, 7, 13, 18, 13, 9, 8, 10, 10, 15, 14, 10, 11, 11, 10])
CUST_N = n + 1

# --- Chi tiết đơn: 1 dòng = 1 đơn ---
det = sorted(orders_all, key=lambda o: o['date'] or '', reverse=True)
wd = wb.create_sheet('Chi tiết đơn')
H2 = ['Ngày', 'Tháng', 'ID khách (email)', 'Phân loại khách', 'Line', 'Mã KH (gốc)', 'Tên khách', 'SĐT', 'Người nhận',
      'Quốc gia nhận', 'Thành phố nhận', 'Mã vận đơn', 'Số kiện', 'Kg', 'Số tiền (VNĐ)', 'Trạng thái', 'Thành phố (quốc gia)', 'Mã nguồn / NV']
sheet(wd, H2, [[o['date'], (o['date'] or '')[:7], o['_cust']['id'], o['_cust']['seg'], o['line'], o['code_raw'],
                 o['name'], o['phone_raw'], o['recv'], o['nation'], o['city'], o['tracking'], o['packages'], o['kg'],
                 o['money'], o['status'], f"{o['city']} ({o['nation']})", src_code(norm_code(o['code_raw'])) or o['_cust']['nv']] for o in det],
      [None, None, None, None, None, '@', None, '@', None, None, None, '@', INT, KG, VND, None, None, None],
      [11, 9, 30, 22, 14, 16, 24, 14, 22, 13, 18, 18, 8, 8, 14, 26, 26, 9])
DN = len(det) + 1
DET = "'Chi tiết đơn'"

def rng(col):
    return f"{DET}!${col}$2:${col}${DN}"

# --- Bảng tổng hợp theo chiều: công thức trên Chi tiết đơn ---
def summary(title, label, key_col, keys, distinct):
    """distinct: {key: (số khách, số khách quay lại)} tính sẵn bằng Python."""
    ws = wb.create_sheet(title)
    H = [label, 'Số đơn', 'Số khách', 'Khách quay lại (≥2 đơn)', 'Tỉ lệ quay lại',
         'Số đơn có kg & tiền', 'Tổng kg', 'Tổng tiền (VNĐ)', '% doanh thu', 'TB / đơn (VNĐ)', 'TB / kg (VNĐ)']
    rows = []
    last = len(keys) + 1
    for i, k in enumerate(keys, 2):
        rows.append([k, f'=COUNTIFS({rng(key_col)},$A{i})', distinct[k][0], distinct[k][1],
                     f'=IF(C{i}>0,D{i}/C{i},0)', f'=COUNTIFS({rng(key_col)},$A{i},{rng("O")},">0")',
                     f'=SUMIFS({rng("N")},{rng(key_col)},$A{i})', f'=SUMIFS({rng("O")},{rng(key_col)},$A{i})',
                     f'=IF(SUM($H$2:$H${last})>0,H{i}/SUM($H$2:$H${last}),0)', f'=IF(F{i}>0,H{i}/F{i},0)', f'=IF(G{i}>0,H{i}/G{i},0)'])
    sheet(ws, H, rows, [None, INT, INT, INT, PCT, INT, KG, VND, PCT, VND, VND], [30, 9, 10, 11, 10, 11, 11, 16, 10, 14, 13])

def build(title, label, col, keyf, limit=None, keys=None):
    per = collections.defaultdict(collections.Counter)
    money = collections.defaultdict(float); cnt = C()
    for o in orders_all:
        k = keyf(o)
        per[k][o['_cust']['id']] += 1
        money[k] += o['money'] or 0; cnt[k] += 1
    if keys is None:
        keys = sorted(per, key=lambda k: (-money[k], -cnt[k]))
    if limit:
        keys = keys[:limit]
    summary(title, label, col, keys, {k: (len(per[k]), sum(1 for v in per[k].values() if v >= 2)) for k in keys})

build('Theo line', 'Line', 'E', lambda o: o['line'])
build('Theo quốc gia', 'Quốc gia nhận', 'J', lambda o: o['nation'])
build('Theo thành phố', 'Thành phố nhận (quốc gia)', 'Q', lambda o: f"{o['city']} ({o['nation']})", limit=500)
build('Theo mã nguồn', 'Mã nguồn / NV', 'R', lambda o: src_code(norm_code(o['code_raw'])) or o['_cust']['nv'] or '(không rõ)')
build('Theo tháng', 'Tháng', 'B', lambda o: (o['date'] or '')[:7],
      keys=sorted({(o['date'] or '')[:7] for o in orders_all if o['date']}))

# --- Theo phân loại / kênh: tổng hợp trên sheet Khách hàng ---
KH = "'Khách hàng'"
def krng(col): return f"{KH}!${col}$2:${col}${CUST_N}"
def cust_summary(title, label, col, keys):
    ws = wb.create_sheet(title)
    H = [label, 'Số khách', 'Khách quay lại (≥2 đơn)', 'Tỉ lệ quay lại', 'Số đơn', 'Tổng kg', 'Tổng tiền (VNĐ)', '% doanh thu']
    last = len(keys) + 1
    rows = [[k, f'=COUNTIFS({krng(col)},$A{i})', f'=COUNTIFS({krng(col)},$A{i},{krng("P")},">=2")', f'=IF(B{i}>0,C{i}/B{i},0)',
             f'=SUMIFS({krng("P")},{krng(col)},$A{i})', f'=SUMIFS({krng("R")},{krng(col)},$A{i})', f'=SUMIFS({krng("S")},{krng(col)},$A{i})',
             f'=IF(SUM($G$2:$G${last})>0,G{i}/SUM($G$2:$G${last}),0)'] for i, k in enumerate(keys, 2)]
    sheet(ws, H, rows, [None, INT, INT, PCT, INT, KG, VND, PCT], [36, 10, 12, 10, 10, 11, 16, 10])

SEGS = ['VIP', 'Thân thiết', 'Giá trị cao (1 đơn)', 'Cần chăm sóc lại (≥2 đơn, >90 ngày)', 'Mới (1 đơn, ≤60 ngày)', 'Một lần, đã lâu (>60 ngày)']
cust_summary('Phân loại khách', 'Phân loại', 'C', SEGS)
cust_summary('Theo kênh', 'Kênh (từ mã KH)', 'N', [k for k, _ in C(c['kenh'] for c in custs if c['kenh']).most_common()])

# --- Ghi chú ---
wn = wb.create_sheet('Ghi chú')
min_d = min(o['date'] for o in orders_all if o['date']); max_d = max(o['date'] for o in orders_all if o['date'])
min_e = min(o['date'] for o in erp); max_e = max(o['date'] for o in erp)
notes = [
    'Ghi chú cách tính',
    f'Nguồn: DB website (mã vận đơn, danh mục mã khách hàng) và DB quản lý đơn (bảng line_orders), bản copy ngày 24/10/2026. Đơn từ {min_d} đến {max_d}: {len(orders_all):,} đơn, {len(custs):,} khách.',
    f'Kg và số tiền chỉ được ghi nhận cho đơn đóng hàng từ {min_e} đến {max_e} (trước đó hệ thống không lưu cân nặng / số tiền). Đơn cũ hơn vẫn được đếm vào Số đơn nhưng ô Kg / Số tiền để trống. Vì vậy "Tổng tiền" và xếp hạng theo tiền phản ánh giai đoạn này; "Số đơn" phản ánh toàn bộ.',
    'Một đơn = một lần gửi: các bản ghi cùng khách, cùng line, cách nhau ≤14 ngày gộp thành 1 đơn (một lần gửi có form đặt đơn, mã vận đơn tạo sau đó và mỗi thùng một mã). Đơn có kg/tiền được gắn vào lần gửi cùng line gần ngày nhất (±21 ngày); nếu không khớp thì là đơn riêng.',
    'Đã loại: line dịch vụ chi nhánh (Tự gửi lên kho, Shopee, Viettel Post lấy hàng — là chặng gom của đơn quốc tế, tính vào sẽ trùng kg), đơn bị đánh dấu sai, đơn test (mã/tên chứa "test", "Nguyễn Văn A").',
    'Số tiền (VNĐ) = doanh thu (số khách trả, với line có khái niệm này) hoặc tổng bill quy đổi VNĐ, cộng khoản khách trả bù nếu có. Kg = kg xác nhận tại kho, nếu không có thì kg khai báo.',
    f'ID khách = email (trường "user" trên mã vận đơn — website cũng tra cứu khách theo email này), viết thường, bỏ khoảng trắng. Đơn không ghi email được gán email của đơn khác cùng mã KH, rồi cùng SĐT (chỉ khi mã/SĐT đó chỉ gắn với đúng 1 email): {C(o["email_src"] for o in allo)["Theo mã KH"]:,} bản ghi theo mã KH, {C(o["email_src"] for o in allo)["Theo SĐT"]:,} theo SĐT. Còn lại {sum(1 for c in custs if not c["has_email"]):,} khách không có email, ID ghi "(không email) + mã KH/SĐT/tên". Đơn trong DB quản lý đơn lấy email người gửi, thiếu thì email người nhận (giống báo cáo ERP). Email của công ty (chứa "amamy", "okhub") không tính là email khách. Mã KH được chuẩn hoá (vd "Amamy DE FLB 228 Thuy Anh" → DEFLB228).',
    'Thành phố nhận lấy từ tỉnh/thành người nhận (đã bỏ mã bưu chính). Line chiều về Việt Nam (Đức→Việt, Nhật→Việt, Mỹ→Việt, Úc→Việt) là tỉnh/thành ở Việt Nam. Quốc gia nhận: theo trường quốc gia trên đơn, nếu trống thì theo điểm đến của line. Tên/SĐT/thành phố của khách = giá trị xuất hiện nhiều nhất.',
    'Kênh lấy từ danh mục mã khách hàng. Mã nguồn / NV = nhóm chữ sau F trong mã KH (vd DEFPT193 → PT), nếu mã không có dạng này thì lấy mã nhân viên trong danh mục mã khách hàng.',
    f'Phân loại: VIP = ≥5 đơn, hoặc ≥2 đơn và tổng tiền thuộc top 10% (≥ {top10:,.0f} VNĐ); Cần chăm sóc lại = ≥2 đơn nhưng >90 ngày chưa gửi; Thân thiết = ≥2 đơn còn lại; Giá trị cao (1 đơn) = 1 đơn nhưng tiền thuộc top 10%; Mới = 1 đơn trong 60 ngày gần đây; Một lần, đã lâu = 1 đơn, >60 ngày.',
    f'Đã đối chiếu export website: {len(wp) + n_test_wp:,} mã vận đơn đang publish trong DB, loại {n_test_wp} đơn test, còn {len(wp):,}.',
    'Sheet Khách hàng sắp theo tổng tiền giảm dần, rồi số đơn. Các sheet "Theo ..." tính bằng công thức COUNTIFS/SUMIFS trên sheet Chi tiết đơn / Khách hàng; riêng cột Số khách & Khách quay lại ở các sheet Theo line/quốc gia/thành phố/mã nguồn/tháng là số tính sẵn. Theo thành phố hiển thị 500 thành phố đứng đầu.',
    'Gợi ý marketing: lọc Phân loại = "Cần chăm sóc lại" để gửi ưu đãi kéo khách quay lại; lọc Số line ≥ 2 để tìm khách dùng nhiều line; sheet Theo thành phố để chọn khu vực chạy quảng cáo.',
]
for t in notes:
    wn.append([t])
wn['A1'].font = Font(name=F, bold=True, size=12)
for (c,) in wn.iter_rows(min_row=2):
    c.font = BODY
    c.alignment = Alignment(wrap_text=True, vertical='top')
wn.column_dimensions['A'].width = 140

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('customers', len(custs), 'orders', len(orders_all), 'paid', sum(1 for o in orders_all if o['money'] is not None), 'erp', len(erp), 'top10', top10)
print(C(c['seg'] for c in custs))
L = collections.defaultdict(lambda: [0, 0, 0.0])
for o in orders_all:
    x = L[o['line']]; x[0] += 1; x[1] += (o['kg'] or 0); x[2] += (o['money'] or 0)
for k, v in sorted(L.items(), key=lambda kv: -kv[1][2]): print(k, v)
for c in custs[:5]: print(c['id'], c['code'], c['name'], c['lines'], c['n_orders'], c['n_paid'], c['kg'], c['money'])
