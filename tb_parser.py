# -*- coding: utf-8 -*-
"""Phân tích tin nhắn thông báo (dán từ Zalo) thành các trường có cấu trúc.

Chạy hoàn toàn bằng luật (regex), không gửi dữ liệu ra ngoài. Kết quả luôn được người dùng
xem lại/sửa trên giao diện trước khi tạo văn bản nên ưu tiên "đọc được nhiều kiểu" hơn "đoán bừa":
trường nào không chắc sẽ để trống và đưa vào danh sách cảnh báo.
"""
import re
import unicodedata
from datetime import date, timedelta

THU = ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy', 'Chủ Nhật']
_THU_KEY = {'hai': 0, 'ba': 1, 'tu': 2, 'nam': 3, 'sau': 4, 'bay': 5}

EMOJI = re.compile('[\U0001F000-\U0001FAFF\u2300-\u23FF\u2600-\u27BF\u2B00-\u2BFF\uFE0F\u200d\u20e3]+')
URL = re.compile(r'https?://[^\s<>"\']+')

# Một giờ: 10:00 | 8h | 8h30 | 15 h 30
_T = r'(\d{1,2})(?::(\d{2})|\s*[hH]\s*(\d{2})?)'
TIME_RANGE = re.compile(_T + r'\s*(?:-|–|—|~|den|toi)\s*' + _T)
TIME_ONE = re.compile(_T)
DATE_FULL = re.compile(r'(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{4}|\d{2})\b')
DATE_SHORT = re.compile(r'(?<![\d:])(\d{1,2})\s*/\s*(\d{1,2})(?![\d/:])')

# (biểu tượng, loại) – ưu tiên nhận theo biểu tượng; không có thì nhận theo từ khóa đầu dòng
EMOJI_KIND = [('🗓', 'time'), ('📅', 'time'), ('📍', 'place'), ('🎁', 'benefit'), ('🚌', 'transport'),
              ('🚍', 'transport'), ('🚐', 'transport'), ('👥', 'quota'), ('✍', 'register'), ('⏰', 'deadline'),
              ('⏳', 'deadline')]
KEYWORDS = [(r'(thoi gian|thoi diem|ngay to chuc)\b', 'time'), (r'(dia diem|noi to chuc)\b', 'place'),
            (r'(quyen loi|ho tro|uu dai)\b', 'benefit'),
            (r'(phuong tien|xe dua don|di chuyen)\b', 'transport'),
            (r'(chi tieu|so luong)\b', 'quota'),
            (r'(cach dang ky|hinh thuc dang ky|dang ky tai|cach thuc dang ky)\b', 'register'),
            (r'(han chot|han dang ky|han cuoi|het han)\b', 'deadline')]
LABEL = re.compile(r'^(thoi gian|thoi diem|ngay to chuc|dia diem|noi to chuc|quyen loi|ho tro|uu dai|phuong tien|'
                   r'chi tieu|so luong|cach dang ky|hinh thuc dang ky|cach thuc dang ky|han chot|han dang ky|han cuoi)'
                   r'\s*[:\-–]\s*')


def fold(s):
    """Bỏ dấu tiếng Việt + hạ chữ thường, GIỮ NGUYÊN độ dài (để dùng chỉ số khớp với chuỗi gốc)."""
    out = []
    for ch in s:
        if ch in 'đĐ':
            out.append('d')
            continue
        dec = unicodedata.normalize('NFD', ch)
        out.append(dec[0] if dec else ch)
    return ''.join(out).lower()


def clean(line):
    s = EMOJI.sub(' ', line)
    s = re.sub(r'^[\s\-\u2022*·▪►>]+', '', s)
    return re.sub(r'\s+', ' ', s).strip()


def _value(c, f):
    """Bỏ nhãn đầu dòng ("Thời gian:", "Địa điểm:"…) nếu có."""
    m = LABEL.match(f)
    return c[m.end():].strip() if m else c


def _hm(m, base=1):
    h = int(m.group(base))
    mi = m.group(base + 1) or m.group(base + 2) or '0'
    mi = int(mi)
    if h > 23 or mi > 59:
        return None
    return '%02d:%02d' % (h, mi)


def _times(f):
    m = TIME_RANGE.search(f)
    if m:
        a, b = _hm(m, 1), _hm(m, 4)
        if a and b:
            return a, b
    m = TIME_ONE.search(f)
    if m:
        a = _hm(m, 1)
        if a:
            return a, ''
    return '', ''


def _mk_date(d, mo, y, today):
    try:
        if y is None:
            y = today.year
            cand = date(y, mo, d)
            if (today - cand).days > 180:  # ghi "05/01" vào cuối năm -> năm sau
                cand = date(y + 1, mo, d)
            return cand
        y = int(y)
        if y < 100:
            y += 2000
        return date(y, mo, d)
    except ValueError:
        return None


def _date(f, today):
    m = DATE_FULL.search(f)
    if m:
        return _mk_date(int(m.group(1)), int(m.group(2)), m.group(3), today)
    m = DATE_SHORT.search(f)
    if m:
        return _mk_date(int(m.group(1)), int(m.group(2)), None, today)
    return None


def _relative(f, today):
    if 'hom nay' in f:
        return today
    if 'ngay mai' in f or 'mai' == f.strip():
        return today + timedelta(days=1)
    if 'ngay kia' in f:
        return today + timedelta(days=2)
    return None


def _kind(raw, f):
    if re.match(r'khoa\s*\d', f):
        return 'khoa'
    for ch, k in EMOJI_KIND:
        if ch in raw:
            return k
    for pat, k in KEYWORDS:
        if re.match(pat, f):
            return k
    return None


def _strip_title(c, f):
    for p in (r'dang ky tham du\s*', r'dang ky\s*', r'thong bao\s*(ve viec|v/v)?\s*[:\-–]?\s*', r'moi tham du\s*'):
        m = re.match(p, f)
        if m and m.end() < len(c):
            return c[m.end():].strip(' :-–'), c[:m.end()].strip().lower()
    return c, ''


def parse(text, today=None):
    """Trả về dict các trường + 'canh_bao' (danh sách chuỗi)."""
    today = today or date.today()
    kq = {'ten_su_kien': '', 'muc_dich': 'đăng ký tham dự', 'ngay': '', 'gio_bd': '', 'gio_kt': '',
          'dia_diem': '', 'quyen_loi': '', 'phuong_tien': '', 'cach_dang_ky': '', 'tong_chi_tieu': 0,
          'chi_tieu': [], 'han_chot_ngay': '', 'han_chot_gio': '', 'link_cu': [], 'canh_bao': []}
    warn = kq['canh_bao']
    text = (text or '').replace('\r', '')
    kq['link_cu'] = [u.rstrip('.,;)') for u in URL.findall(text)]
    lines = [(l, clean(l)) for l in text.split('\n')]
    lines = [(r, c) for r, c in lines if c]
    if not lines:
        return kq

    # --- tiêu đề: dòng đầu tiên ---
    raw0, c0 = lines[0]
    if _kind(raw0, fold(c0)) in (None,):
        name, pre = _strip_title(c0, fold(c0))
        kq['ten_su_kien'] = name
        if 'dang ky' in fold(pre):
            kq['muc_dich'] = 'đăng ký tham dự'
        elif pre:
            kq['muc_dich'] = pre.strip(' :')
        lines = lines[1:]

    declared = 0
    for raw, c in lines:
        f = fold(c)
        k = _kind(raw, f)
        if k is None:
            continue
        v = _value(c, f)
        fv = fold(v)
        if k == 'time':
            a, b = _times(fv)
            kq['gio_bd'], kq['gio_kt'] = kq['gio_bd'] or a, kq['gio_kt'] or b
            d = _date(fv, today)
            if d and not kq['ngay']:
                kq['ngay'] = d.isoformat()
                mt = re.search(r'thu\s*(hai|ba|tu|nam|sau|bay)\b|chu\s*nhat', fv)
                if mt:
                    idx = 6 if 'chu' in mt.group(0) else _THU_KEY[mt.group(1)]
                    if idx != d.weekday():
                        warn.append('Thông báo ghi "%s" nhưng ngày %s là %s – kiểm tra lại ngày.'
                                    % (THU[idx], d.strftime('%d/%m/%Y'), THU[d.weekday()]))
        elif k == 'place':
            kq['dia_diem'] = kq['dia_diem'] or v
        elif k == 'benefit':
            kq['quyen_loi'] = kq['quyen_loi'] or v
        elif k == 'transport':
            kq['phuong_tien'] = kq['phuong_tien'] or v
        elif k == 'quota':
            m = re.search(r'(\d+)', fv)
            if m:
                declared = int(m.group(1))
        elif k == 'khoa':
            for m in re.finditer(r'khoa\s*(\d{1,3})\s*[:\-–]?\s*(\d+)\s*(?:em|sv|sinh vien|ban|suat|cho|nguoi)?', f):
                note = ''
                mn = re.match(r'\s*\((.*?)\)', c[m.end():])
                if mn:
                    note = mn.group(1).strip()
                kq['chi_tieu'].append({'khoa': m.group(1), 'so_luong': int(m.group(2)), 'ghi_chu': note})
        elif k == 'register':
            v2 = re.sub(r'\s*(tai|theo|qua|o)?\s*(link|duong dan|lien ket)\s*:?\s*$', '', fv)
            kq['cach_dang_ky'] = kq['cach_dang_ky'] or v[:len(v2)].strip(' :')
        elif k == 'deadline':
            a, _ = _times(fv)
            d = _relative(fv, today) or _date(fv, today)
            kq['han_chot_gio'] = kq['han_chot_gio'] or a
            if d and not kq['han_chot_ngay']:
                kq['han_chot_ngay'] = d.isoformat()

    tong = sum(x['so_luong'] for x in kq['chi_tieu'])
    kq['tong_chi_tieu'] = declared or tong
    if declared and tong and declared != tong:
        warn.append('Tổng chỉ tiêu ghi %d nhưng cộng các khóa được %d – kiểm tra lại số lượng.' % (declared, tong))
    if not kq['chi_tieu']:
        warn.append('Chưa nhận ra chỉ tiêu theo khóa (dạng "Khóa 24: 12 em") – hãy nhập ở bảng bên cạnh.')
    for key, nhan in (('ten_su_kien', 'tên sự kiện'), ('ngay', 'ngày diễn ra'), ('dia_diem', 'địa điểm')):
        if not kq[key]:
            warn.append('Chưa tìm thấy %s trong tin nhắn.' % nhan)
    if (kq['han_chot_gio'] and not kq['han_chot_ngay']):
        warn.append('Có giờ hạn chót nhưng chưa rõ ngày – hãy chọn ngày hạn chót.')
    return kq
