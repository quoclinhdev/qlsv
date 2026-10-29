# -*- coding: utf-8 -*-
"""Đọc file PDF (văn bản hành chính) và phân tích thành các trường.

Quy trình:
  1. Lấy chữ có sẵn trong PDF (PyMuPDF).
  2. Nếu PDF là bản scan, hoặc chữ bị hỏng font tiếng Việt (copy ra "BÖ GlÅo DVC VÄ DAO TAO") -> OCR tiếng Việt bằng Tesseract.
  3. Phân tích: loại văn bản, số hiệu, ngày ban hành, tên/trích yếu, thời gian, địa điểm, tổ chức thực hiện.

Cài đặt:
  pip install pymupdf pytesseract pillow
  Cài thêm phần mềm Tesseract OCR (Windows: bản của UB-Mannheim) và NHỚ chọn gói ngôn ngữ "Vietnamese".
  Nếu Tesseract không nằm trong PATH: đặt biến môi trường TESSERACT_CMD = C:\\Program Files\\Tesseract-OCR\\tesseract.exe
  (tuỳ chọn) TB_OCR_LANG: mặc định 'vie'.
"""
import io
import os
import re
from datetime import date

from tb_parser import fold

MAX_TRANG = 40
VN_CHARS = set('àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ')
VN_CHARS |= {c.upper() for c in VN_CHARS}
# ký tự hay xuất hiện khi font tiếng Việt cũ bị giải mã sai, không có trong tiếng Việt chuẩn
SUS_CHARS = set('ÅÄÖÜåäöüÆæØøÞþßÐð¥£§¤¦¨©ª«¬®¯°±²³´¶¸¹º»¼½¾¿×÷')


class LoiPDF(RuntimeError):
    """Lỗi có thể giải thích cho người dùng (thiếu thư viện, thiếu OCR, file hỏng…)."""


def bi_hong_font(text):
    """True nếu chữ trích ra từ PDF trông như tiếng Việt bị mất dấu / sai mã."""
    letters = [c for c in text if c.isalpha()]
    if len(letters) < 100:
        return False
    vn = sum(c in VN_CHARS for c in letters)
    sus = sum(c in SUS_CHARS for c in letters)
    return vn / len(letters) < 0.05 or sus / len(letters) > 0.02


def _ocr(doc):
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        raise LoiPDF('PDF này cần đọc bằng OCR nhưng chưa cài thư viện: pip install pytesseract pillow '
                     '(và cài phần mềm Tesseract OCR kèm ngôn ngữ Vietnamese).')
    cmd = (os.environ.get('TESSERACT_CMD') or '').strip().strip('"')
    if not cmd:   # tự dò thư mục cài mặc định (Windows) khi chưa đặt biến môi trường
        import shutil
        cmd = shutil.which('tesseract') or next((p for p in (
            r'C:\Program Files\Tesseract-OCR\tesseract.exe', r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            os.path.expandvars(r'%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe')) if os.path.exists(p)), '')
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd
    lang = os.environ.get('TB_OCR_LANG', 'vie')
    try:
        co = pytesseract.get_languages(config='')
    except Exception as e:
        raise LoiPDF('Không chạy được Tesseract OCR (đường dẫn đang dùng: %s). Chi tiết: %s. Hãy cài Tesseract (kèm gói '
                     'Vietnamese) và đặt biến môi trường TESSERACT_CMD trỏ tới tesseract.exe nếu nó chưa nằm trong PATH.'
                     % (pytesseract.pytesseract.tesseract_cmd, str(e).strip()[:200]))
    if lang not in co:
        raise LoiPDF('Tesseract chưa có ngôn ngữ "%s". Cài lại Tesseract và chọn gói "Vietnamese" '
                     '(hoặc chép file vie.traineddata vào thư mục tessdata).' % lang)
    out = []
    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes('png')))
        out.append(pytesseract.image_to_string(img, lang=lang))
    return '\n'.join(out)


def doc_pdf(data):
    """bytes PDF -> (văn_bản, phương_pháp, [cảnh_báo]). phương_pháp: 'chu' (có sẵn chữ) hoặc 'ocr'."""
    if not data[:5] == b'%PDF-':
        raise LoiPDF('File tải lên không phải PDF.')
    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz   # PyMuPDF bản cũ
        except ImportError:
            raise LoiPDF('Chưa cài thư viện đọc PDF: pip install pymupdf')
    try:
        doc = fitz.open(stream=data, filetype='pdf')
    except Exception:
        raise LoiPDF('Không mở được file PDF (file hỏng hoặc có mật khẩu).')
    if doc.needs_pass:
        raise LoiPDF('PDF có mật khẩu, hãy bỏ mật khẩu rồi tải lại.')
    if doc.page_count > MAX_TRANG:
        raise LoiPDF('PDF quá dài (%d trang), tối đa %d trang.' % (doc.page_count, MAX_TRANG))
    warn = []
    text = '\n'.join(p.get_text('text') for p in doc)
    if len(text.strip()) < 40:
        warn.append('PDF là bản scan (không có chữ) – đã đọc bằng OCR, hãy kiểm tra lại kết quả.')
        return _ocr(doc), 'ocr', warn
    if bi_hong_font(text):
        warn.append('Chữ trong PDF bị lỗi font tiếng Việt – đã đọc lại bằng OCR, hãy kiểm tra lại kết quả.')
        return _ocr(doc), 'ocr', warn
    return text, 'chu', warn


# ------------------------------------------------------------------ phân tích văn bản
LOAI = ['ke hoach', 'thong bao', 'quyet dinh', 'cong van', 'to trinh', 'bao cao', 'chuong trinh', 'huong dan',
        'ke hoach tong the']
ROMAN = re.compile(r'^(i{1,3}|iv|v|vi{1,3}|ix|x)\s*[\.\)]\s*\S')
BULLET = re.compile(r'^[-\u2013\u2014\u2022\u00b7\u25aa\u25ba*]\s*')
_T = r'(\d{1,2})(?::(\d{2})|\s*h\s*(\d{2})?)(?![a-z])'
TIME_RANGE = re.compile(_T + r'\s*(?:-|–|—|~|den)\s*' + _T)
TIME_ONE = re.compile(_T)
RANGE_1 = re.compile(r'(\d{1,2})\s*(?:-|–|—|den)\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})')
RANGE_2 = re.compile(r'(\d{1,2})\s*/\s*(\d{1,2})\s*(?:-|–|—|den)\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})')
ONE = re.compile(r'(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})')


def _lines(text):
    out = []
    for l in text.replace('\r', '').split('\n'):
        l = re.sub(r'\s+', ' ', l).strip()
        if not l or re.fullmatch(r'[-– ]*\d{1,3}[-– ]*', l):   # bỏ số trang đứng một mình
            continue
        out.append(l)
    return out


def _mk(d, m, y):
    try:
        return date(int(y), int(m), int(d))
    except ValueError:
        return None


def _gio(f):
    m = TIME_RANGE.search(f)
    if m:
        a, b = _hm(m, 1), _hm(m, 4)
        if a and b:
            return a, b
    m = TIME_ONE.search(f)
    if m and _hm(m, 1):
        return _hm(m, 1), ''
    return '', ''


def _hm(m, base):
    h, mi = int(m.group(base)), int(m.group(base + 1) or m.group(base + 2) or 0)
    return '%02d:%02d' % (h, mi) if h < 24 and mi < 60 else None


def _sau_nhan(raw, f, pat):
    """Tìm nhãn bằng regex trên bản đã bỏ dấu; trả về phần giá trị sau nhãn (từ chuỗi gốc) hoặc None."""
    m = re.match(pat, f)
    return raw[m.end():].strip(' :-–') if m else None


def _bo_tien_to(s, pats):
    f = fold(s)
    for p in pats:
        m = re.match(p, f)
        if m and m.end() < len(s):
            return s[m.end():].strip(' :-–')
    return s


def phan_tich_van_ban(text, today=None):
    today = today or date.today()
    kq = {'ten_su_kien': '', 'muc_dich': 'tham dự', 'ngay': '', 'ngay_kt': '', 'gio_bd': '', 'gio_kt': '',
          'thoi_gian_text': '', 'dia_diem': '', 'so_hieu': '', 'loai_vb': '', 'ngay_vb': '', 'quyen_loi': '',
          'phuong_tien': '', 'cach_dang_ky': '', 'han_chot_ngay': '', 'han_chot_gio': '', 'chi_tieu': [],
          'tong_chi_tieu': 0, 'link_cu': [], 'to_chuc': [], 'cau_truc': [], 'canh_bao': []}
    warn = kq['canh_bao']
    lines = _lines(text)
    if not lines:
        warn.append('Không đọc được chữ nào từ PDF.')
        return kq
    F = [fold(l) for l in lines]
    head = '\n'.join(lines[:30])

    m = re.search(r'(\d{1,5})\s*/\s*([A-ZĐ][A-ZĐ0-9\-]{1,})', head)
    if m:
        kq['so_hieu'] = '%s/%s' % (m.group(1), m.group(2))
    m = re.search(r'ngay\s*(\d{1,2})\s*thang\s*(\d{1,2})\s*nam\s*(\d{4})', fold(head))
    if m and _mk(*m.groups()):
        kq['ngay_vb'] = _mk(*m.groups()).isoformat()

    idx = None
    for i in range(min(40, len(lines))):
        if F[i].strip() in LOAI:
            kq['loai_vb'], idx = lines[i].strip().upper(), i
            break
    if idx is not None:
        tieu = []
        for j in range(idx + 1, min(idx + 6, len(lines))):
            if re.match(r'(can cu|thuc hien|nham |theo |kinh gui|\d+\s*[\.\)])', F[j]) or ROMAN.match(F[j]):
                break
            tieu.append(lines[j])
        ten = ' '.join(tieu)
        ten = _bo_tien_to(ten, [r'(ve viec|v/v)\s*[:\-–]?\s*'])
        ten = _bo_tien_to(ten, [r'to chuc\s+'])
        kq['ten_su_kien'] = ten.strip()
    else:
        warn.append('Không tìm thấy tên loại văn bản (KẾ HOẠCH / THÔNG BÁO…) – hãy nhập tên sự kiện thủ công.')
    if not kq['ten_su_kien'] and idx is not None:
        warn.append('Không nhận ra trích yếu/tên sự kiện – hãy nhập thủ công.')

    # --- thời gian / địa điểm ---
    for i, f in enumerate(F):
        v = _sau_nhan(lines[i], f, r'^(\d+\s*[\.\)]\s*)?(thoi gian|thoi diem)(\s*to chuc)?\s*[:\-–]?\s*')
        if v is not None and not kq['thoi_gian_text']:
            if not v and i + 1 < len(lines):
                v = lines[i + 1]
            kq['thoi_gian_text'] = v[:300]
            fv = fold(v)
            m1, m2, m3 = RANGE_2.search(fv), RANGE_1.search(fv), ONE.search(fv)
            if m1:
                a, b = _mk(m1.group(1), m1.group(2), m1.group(5)), _mk(m1.group(3), m1.group(4), m1.group(5))
            elif m2:
                a, b = _mk(m2.group(1), m2.group(3), m2.group(4)), _mk(m2.group(2), m2.group(3), m2.group(4))
            elif m3:
                a, b = _mk(m3.group(1), m3.group(2), m3.group(3)), None
            else:
                a = b = None
            if a:
                kq['ngay'], kq['ngay_kt'] = a.isoformat(), (b.isoformat() if b and b != a else '')
            kq['gio_bd'], kq['gio_kt'] = _gio(fv)
        v = _sau_nhan(lines[i], f, r'^(\d+\s*[\.\)]\s*)?(dia diem|noi to chuc)\s*[:\-–]?\s*')
        if v is not None and not kq['dia_diem']:
            cho = [v] if v else []
            for j in range(i + 1, len(lines)):
                if ROMAN.match(F[j]) or re.match(r'(\d+\s*[\.\)]\s|to chuc thuc hien|noi nhan)', F[j]):
                    break
                cho.append(BULLET.sub('', lines[j]))
                if len(cho) >= 6:
                    break
            kq['dia_diem'] = '; '.join(x for x in cho if x)[:300]
    if not kq['thoi_gian_text']:
        warn.append('Không thấy mục "Thời gian:" – hãy nhập thời gian thủ công.')
    elif not kq['ngay']:
        warn.append('Có mục thời gian nhưng chưa tách được ngày cụ thể – nội dung được giữ nguyên văn.')
    if not kq['dia_diem']:
        warn.append('Không thấy mục "Địa điểm:" – hãy nhập thủ công.')

    # --- cấu trúc các phần La Mã ---
    kq['cau_truc'] = [l for l, f in zip(lines, F) if ROMAN.match(f) and len(l) < 90]

    # --- tổ chức thực hiện ---
    start = next((i for i, f in enumerate(F) if 'to chuc thuc hien' in f and len(f) < 40), None)
    if start is not None:
        cur = None
        for l, f in zip(lines[start + 1:], F[start + 1:]):
            if f.startswith('noi nhan'):
                break
            m = re.match(r'^(\d{1,2})\s*[\.\)]\s+(.+)$', l)
            if m:
                cur = {'don_vi': m.group(2).strip(), 'nhiem_vu': []}
                kq['to_chuc'].append(cur)
            elif cur is not None:
                if BULLET.match(l) or not cur['nhiem_vu']:
                    cur['nhiem_vu'].append(BULLET.sub('', l))
                else:
                    cur['nhiem_vu'][-1] += ' ' + l
    return kq