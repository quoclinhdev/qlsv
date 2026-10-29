# -*- coding: utf-8 -*-
"""Soạn nội dung: thông báo đăng Zalo, văn bản hành chính (bản xem trước) và bố cục bảng đăng ký."""
from datetime import date, timedelta

from tb_parser import THU

# >>> Giá trị mặc định cho phần đầu văn bản – sửa cho đúng đơn vị của bạn (cũng sửa được ngay trên giao diện). <<<
DEFAULTS = {
    'don_vi_chu_quan': 'TRƯỜNG ĐẠI HỌC CỬU LONG',
    'don_vi': 'KHOA CÔNG NGHỆ THÔNG TIN',
    'don_vi_viet': 'Khoa Công nghệ thông tin',   # dùng trong câu văn: "... thông báo đến sinh viên ..."
    'so_vb': '      /TB-CNTT',
    'dia_danh': 'Vĩnh Long',
    'chuc_danh_ky': 'TRƯỞNG KHOA',
    'nguoi_ky': '',
}
TEXT_KEYS = ['ten_su_kien', 'muc_dich', 'dia_diem', 'quyen_loi', 'phuong_tien', 'cach_dang_ky', 'link_sheet',
             'don_vi_chu_quan', 'don_vi', 'don_vi_viet', 'so_vb', 'dia_danh', 'chuc_danh_ky', 'nguoi_ky', 'thoi_gian_text']
DATE_KEYS = ['ngay', 'han_chot_ngay', 'ngay_ban_hanh']
TIME_KEYS = ['gio_bd', 'gio_kt', 'han_chot_gio']
CACH_DK_MAC_DINH = 'điền MSSV và HỌ TÊN vào một dòng trống đúng KHÓA của mình'
TOI_DA_TONG = 1000


def _d(s):
    try:
        return date.fromisoformat(s) if s else None
    except (TypeError, ValueError):
        return None


def fmt_ngay(d):
    return d.strftime('%d/%m/%Y') if d else ''


def ngay_chu(d):
    return 'ngày %02d tháng %02d năm %d' % (d.day, d.month, d.year) if d else 'ngày ... tháng ... năm ...'


def thu_vn(d):
    return THU[d.weekday()] if d else ''


def gio_h(t):
    """'15:30' -> '15h30'; '15:00' -> '15h'."""
    if not t:
        return ''
    h, m = t.split(':')
    return '%dh%s' % (int(h), '' if m == '00' else m)


def _time_ok(t):
    try:
        h, m = t.split(':')
        return 0 <= int(h) <= 23 and 0 <= int(m) <= 59
    except (ValueError, AttributeError):
        return False


def chuan_hoa(data):
    """Làm sạch dữ liệu từ trình duyệt: cắt độ dài, ép kiểu, bỏ dòng chỉ tiêu không hợp lệ."""
    data = data if isinstance(data, dict) else {}
    d = dict(DEFAULTS)
    for k in TEXT_KEYS:
        v = data.get(k)
        if isinstance(v, str) and (v.strip() or k in DEFAULTS):
            d[k] = v.strip()[:300]
        else:
            d.setdefault(k, '')
    d['muc_dich'] = d.get('muc_dich') or 'đăng ký tham dự'
    for k in DATE_KEYS:
        v = data.get(k)
        d[k] = v if isinstance(v, str) and _d(v) else ''
    if not d['ngay_ban_hanh']:
        d['ngay_ban_hanh'] = date.today().isoformat()
    for k in TIME_KEYS:
        v = data.get(k)
        d[k] = v if isinstance(v, str) and _time_ok(v) else ''
    ct, seen = [], set()
    for it in (data.get('chi_tieu') or [])[:12]:
        if not isinstance(it, dict):
            continue
        khoa = str(it.get('khoa', '')).strip()[:8]
        try:
            n = int(it.get('so_luong', 0))
        except (TypeError, ValueError):
            n = 0
        if not khoa or n <= 0 or khoa in seen:
            continue
        seen.add(khoa)
        ct.append({'khoa': khoa, 'so_luong': min(n, 500), 'ghi_chu': str(it.get('ghi_chu', '')).strip()[:200]})
    d['chi_tieu'] = ct
    d['tong_chi_tieu'] = min(sum(x['so_luong'] for x in ct), TOI_DA_TONG)
    return d


def liet_ke(items):
    items = list(items)
    if len(items) <= 1:
        return ''.join(items)
    return ', '.join(items[:-1]) + ' và ' + items[-1]


def khoang_gio(d):
    if d['gio_bd'] and d['gio_kt']:
        return '%s–%s' % (d['gio_bd'], d['gio_kt'])
    return d['gio_bd']


def thoi_gian_zalo(d):
    """Ưu tiên thời gian ghi nguyên văn (vd. 'Từ ngày 10 – 14/8/2026'), không có thì ghép từ giờ + ngày."""
    if d.get('thoi_gian_text'):
        return d['thoi_gian_text']
    n = _d(d['ngay'])
    return ', '.join(x for x in (khoang_gio(d), ('%s %s' % (thu_vn(n), fmt_ngay(n))).strip()) if x)


def loi_ket(d):
    if d['chi_tieu']:
        return 'Đề nghị sinh viên các khóa nêu trên đăng ký đúng khóa của mình và đúng thời hạn. Trân trọng thông báo./.'
    return 'Đề nghị sinh viên nắm thông tin và thực hiện nghiêm túc. Trân trọng thông báo./.'


def co_dang_ky(d):
    return bool(d['chi_tieu'] or d['link_sheet'] or d['cach_dang_ky'])


def han_chot_zalo(d, today):
    g, n = gio_h(d['han_chot_gio']), _d(d['han_chot_ngay'])
    if not g and not n:
        return ''
    if n == today:
        p = 'hôm nay'
    elif n == today + timedelta(days=1):
        p = 'ngày mai'
    elif n:
        p = '%s %s' % (thu_vn(n), fmt_ngay(n))
    else:
        p = ''
    return ' '.join(x for x in (g, p) if x)


def han_chot_hc(d):
    g, n = gio_h(d['han_chot_gio']), _d(d['han_chot_ngay'])
    if not g and not n:
        return ''
    s = 'trước ' + g if g else 'trước'
    return s + (', %s, %s' % (thu_vn(n).lower(), ngay_chu(n)) if n else '')


def dong_chi_tieu(it):
    s = '%s: %d em' % ('Khóa ' + it['khoa'], it['so_luong'])
    return s + (' (%s)' % it['ghi_chu'] if it['ghi_chu'] else '')


# ------------------------------------------------------------------ Thông báo đăng Zalo
def van_ban_zalo(d, today=None):
    today = today or date.today()
    n = _d(d['ngay'])
    L = ['📢 %s %s' % (d['muc_dich'].upper(), d['ten_su_kien'])]
    tg = thoi_gian_zalo(d)
    if tg:
        L.append('🗓 ' + tg)
    if d['dia_diem']:
        L.append('📍 ' + d['dia_diem'])
    if d['quyen_loi']:
        L.append('🎁 ' + d['quyen_loi'])
    if d['phuong_tien']:
        L.append('🚌 ' + d['phuong_tien'])
    if d['chi_tieu']:
        L += ['', '👥 Chỉ tiêu %d em:' % d['tong_chi_tieu']] + [dong_chi_tieu(x) for x in d['chi_tieu']]
    if co_dang_ky(d):
        L += ['', '✍️ Cách đăng ký: %s tại link:' % (d['cach_dang_ky'] or CACH_DK_MAC_DINH),
              '👉 ' + (d['link_sheet'] or '[chưa có link – bấm "Tạo Google Sheet" để lấy link]')]
    hc = han_chot_zalo(d, today)
    if hc:
        L += ['', '⏰ Hạn chót: %s. @All' % hc]
    return '\n'.join(L)


# ------------------------------------------------------------------ Văn bản hành chính
def kinh_gui(d):
    return ('Sinh viên Khóa ' + liet_ke([x['khoa'] for x in d['chi_tieu']])) if d['chi_tieu'] else 'Sinh viên'


def tieu_de_trich_yeu(d):
    return ('Về việc %s %s' % (d['muc_dich'], d['ten_su_kien'])).strip()


def muc_noi_dung(d):
    """Danh sách (nhãn, nội dung) đánh số 1., 2., … dùng chung cho bản xem trước và Word."""
    n = _d(d['ngay'])
    m = []
    tg = khoang_gio(d)
    if d.get('thoi_gian_text'):
        m.append(('Thời gian', d['thoi_gian_text']))
    elif tg or n:
        s = (tg + ', ' if tg else '') + ('%s, %s' % (thu_vn(n).lower(), ngay_chu(n)) if n else '')
        m.append(('Thời gian', s))
    if d['dia_diem']:
        m.append(('Địa điểm', d['dia_diem']))
    if d['quyen_loi']:
        m.append(('Quyền lợi', d['quyen_loi']))
    if d['phuong_tien']:
        m.append(('Phương tiện', d['phuong_tien']))
    if d['chi_tieu']:
        m.append(('Chỉ tiêu', '%d sinh viên, phân bổ như sau:' % d['tong_chi_tieu']))
    if co_dang_ky(d):
        dk = d['cach_dang_ky'] or CACH_DK_MAC_DINH
        dk = dk[:1].upper() + dk[1:]
        m.append(('Cách thức đăng ký', dk + (' tại đường dẫn: ' + d['link_sheet'] if d['link_sheet'] else '') + '.'))
    h = han_chot_hc(d)
    if h:
        m.append(('Thời hạn đăng ký', h[:1].upper() + h[1:] + '.'))
    return m


def van_ban_hanh_chinh(d):
    ngay_bh = _d(d['ngay_ban_hanh'])
    W = 38
    L = [d['don_vi_chu_quan'].ljust(W) + 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',
         d['don_vi'].ljust(W) + 'Độc lập – Tự do – Hạnh phúc',
         ('Số: ' + d['so_vb']).ljust(W) + '%s, %s' % (d['dia_danh'], ngay_chu(ngay_bh)),
         '', '                              THÔNG BÁO', '          ' + tieu_de_trich_yeu(d), '',
         'Kính gửi: ' + kinh_gui(d), '',
         '%s thông báo đến sinh viên về việc %s %s như sau:' % (d['don_vi_viet'], d['muc_dich'], d['ten_su_kien'])]
    for i, (nhan, nd) in enumerate(muc_noi_dung(d), 1):
        L.append('%d. %s: %s' % (i, nhan, nd))
        if nhan == 'Chỉ tiêu':
            L.append('   %-10s%-18s%s' % ('Khóa', 'Số lượng', 'Ghi chú'))
            for it in d['chi_tieu']:
                L.append('   %-10s%-18d%s' % (it['khoa'], it['so_luong'], it['ghi_chu']))
    L += ['', loi_ket(d),
          '', 'Nơi nhận:'.ljust(W) + d['chuc_danh_ky'], '- Như trên;', '- Lưu: VT.'.ljust(W) + d['nguoi_ky']]
    return '\n'.join(L)


# ------------------------------------------------------------------ Bố cục bảng đăng ký (dùng cho Google Sheet và Excel)
HEADER = ['STT', 'Khóa', 'MSSV', 'Họ và tên']
SHEET_TITLE = 'Danh sách'
HEADER_ROW = 4  # dòng 1-3: tiêu đề/thông tin/hướng dẫn, dòng 4: tiêu đề cột (đánh số từ 1)


def bo_cuc(d):
    n = _d(d['ngay'])
    title = ('DANH SÁCH %s %s' % (d['muc_dich'], d['ten_su_kien'])).upper()
    info = []
    tg = thoi_gian_zalo(d)
    if tg:
        info.append('Thời gian: ' + tg)
    if d['dia_diem']:
        info.append('Địa điểm: ' + d['dia_diem'])
    if d['phuong_tien']:
        info.append('Phương tiện: ' + d['phuong_tien'])
    h = han_chot_hc(d)
    if h:
        info.append('Hạn chót đăng ký: ' + h.replace('trước ', '', 1))
    guide = ('CÁCH ĐIỀN: nhập MSSV và HỌ TÊN vào một dòng trống đúng KHÓA của mình (không sửa cột STT, Khóa). '
             'Ví dụ: MSSV 24xxxxxxx – Nguyễn Văn A. MSSV trùng sẽ được tô đỏ.')
    rows, r, stt = [], HEADER_ROW + 1, 0
    for it in d['chi_tieu']:
        label = 'KHÓA %s – %d chỉ tiêu' % (it['khoa'], it['so_luong']) + (' (%s)' % it['ghi_chu'] if it['ghi_chu'] else '')
        rows.append({'kind': 'group', 'row': r, 'label': label})
        r += 1
        for _ in range(it['so_luong']):
            stt += 1
            rows.append({'kind': 'slot', 'row': r, 'stt': stt, 'khoa': it['khoa']})
            r += 1
    return {'title': title, 'info': '\n'.join(info), 'info_lines': max(1, len(info)), 'guide': guide,
            'header': HEADER, 'rows': rows, 'first': HEADER_ROW + 1, 'last': r - 1, 'total': stt}