# -*- coding: utf-8 -*-
"""Tạo Google Sheet đăng ký từ bố cục `bo_cuc` và cấp quyền cho sinh viên.

Cần tài khoản dịch vụ (service account) của Google Cloud, bật 2 API: Google Sheets API và Google Drive API.
Cấu hình bằng biến môi trường:
    GOOGLE_SA_FILE     đường dẫn file JSON khóa của service account   (hoặc)
    GOOGLE_SA_JSON     nội dung JSON của khóa
    TB_OWNER_EMAIL     (tuỳ chọn) email của bạn – sheet sẽ được chia sẻ quyền chỉnh sửa cho email này
                       để bạn thấy nó trong Google Drive (file do service account làm chủ sở hữu).
    TB_ANYONE_ROLE     (tuỳ chọn) 'writer' (mặc định: ai có link đều sửa được) hoặc 'reader'

Sinh viên chỉ sửa được cột MSSV và Họ tên của các dòng chỉ tiêu; phần tiêu đề, cột STT/Khóa bị khóa.
"""
import json
import os

from tb_van_ban import SHEET_TITLE, bo_cuc

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
NAVY, BLUE, PALE = (0.118, 0.227, 0.541), (0.145, 0.388, 0.922), (0.863, 0.902, 0.980)
WHITE = (1, 1, 1)


class ChuaCauHinh(RuntimeError):
    """Chưa có khóa Google – giao diện sẽ báo hướng dẫn thay vì lỗi khó hiểu."""


def _dich_vu():
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError:
        raise ChuaCauHinh('Chưa cài thư viện: pip install google-api-python-client google-auth')
    raw, path = os.environ.get('GOOGLE_SA_JSON'), os.environ.get('GOOGLE_SA_FILE')
    if raw:
        cred = service_account.Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    elif path and os.path.exists(path):
        cred = service_account.Credentials.from_service_account_file(path, scopes=SCOPES)
    else:
        raise ChuaCauHinh('Chưa cấu hình khóa Google: đặt biến môi trường GOOGLE_SA_FILE (đường dẫn file JSON '
                          'của service account) hoặc GOOGLE_SA_JSON.')
    return (build('sheets', 'v4', credentials=cred, cache_discovery=False),
            build('drive', 'v3', credentials=cred, cache_discovery=False))


class LoiBuoc(RuntimeError):
    """Lỗi Google ở một bước cụ thể (để báo người dùng biết hỏng ở đâu)."""
    def __init__(self, buoc, goc):
        super().__init__(buoc)
        self.buoc, self.goc = buoc, goc


def _chay(buoc, req):
    try:
        return req.execute()
    except Exception as e:
        raise LoiBuoc(buoc, e)


def email_dich_vu():
    """Email của service account (để chia sẻ bảng có sẵn cho nó); '' nếu không đọc được."""
    try:
        raw, path = os.environ.get('GOOGLE_SA_JSON'), os.environ.get('GOOGLE_SA_FILE')
        if raw:
            return json.loads(raw).get('client_email', '')
        if path and os.path.exists(path):
            with open(path, encoding='utf-8') as f:
                return json.load(f).get('client_email', '')
    except Exception:
        pass
    return ''


def _id_tu_link(s):
    import re
    m = re.search(r'/spreadsheets/d/([A-Za-z0-9_\-]{20,})', s or '') or re.fullmatch(r'([A-Za-z0-9_\-]{20,})', (s or '').strip())
    if not m:
        raise ChuaCauHinh('Link bảng có sẵn không đúng dạng https://docs.google.com/spreadsheets/d/…')
    return m.group(1)


def _mau(rgb):
    return {'red': rgb[0], 'green': rgb[1], 'blue': rgb[2]}


def _range(sid, r1, r2, c1=0, c2=4):
    """Chỉ số dòng/cột theo kiểu 0-based, nửa mở [r1, r2)."""
    return {'sheetId': sid, 'startRowIndex': r1, 'endRowIndex': r2, 'startColumnIndex': c1, 'endColumnIndex': c2}


def _fmt(sid, r1, r2, c1, c2, **cell):
    fields = ','.join('userEnteredFormat.' + k for k in cell)
    return {'repeatCell': {'range': _range(sid, r1, r2, c1, c2), 'cell': {'userEnteredFormat': cell}, 'fields': fields}}


def _khoi_slot(rows):
    """Gom các dòng 'slot' liên tiếp theo khóa -> [(dòng_đầu, dòng_cuối)] (đánh số từ 1)."""
    khoi, cur = [], None
    for it in rows:
        if it['kind'] != 'slot':
            cur = None
            continue
        if cur and it['row'] == cur[1] + 1:
            cur[1] = it['row']
        else:
            cur = [it['row'], it['row']]
            khoi.append(cur)
    return khoi


def tao_google_sheet(d, link_co_san=''):
    """Tạo sheet (hoặc điền vào bảng trống có sẵn nếu có link_co_san), định dạng, khóa ô, mở quyền.
    Trả về {'id', 'link', 'chia_se_chu', 'canh_bao'}."""
    sheets, drive = _dich_vu()
    L = bo_cuc(d)
    ten = 'Đăng ký – ' + d['ten_su_kien'][:80]
    luoi = {'frozenRowCount': 4, 'rowCount': max(L['last'] + 5, 30), 'columnCount': 4}
    canh_bao = []
    if link_co_san:
        ss_id = _id_tu_link(link_co_san)
        meta = _chay('mở bảng có sẵn (đã chia sẻ quyền Người chỉnh sửa cho email service account chưa?)',
                     sheets.spreadsheets().get(spreadsheetId=ss_id, fields='sheets.properties.sheetId'))
        sid = meta['sheets'][0]['properties']['sheetId']
        _chay('chuẩn bị bảng có sẵn', sheets.spreadsheets().batchUpdate(spreadsheetId=ss_id, body={'requests': [
            {'updateSpreadsheetProperties': {'properties': {'title': ten}, 'fields': 'title'}},
            {'updateSheetProperties': {'properties': {'sheetId': sid, 'title': SHEET_TITLE, 'gridProperties': luoi},
                                       'fields': 'title,gridProperties(frozenRowCount,rowCount,columnCount)'}}]}))
        _chay('xóa nội dung cũ', sheets.spreadsheets().values().clear(spreadsheetId=ss_id, range="'%s'" % SHEET_TITLE, body={}))
    else:
        f = _chay('tạo bảng tính mới', sheets.spreadsheets().create(body={
            'properties': {'title': ten}, 'sheets': [{'properties': {'title': SHEET_TITLE, 'gridProperties': luoi}}]},
            fields='spreadsheetId,sheets.properties.sheetId'))
        ss_id, sid = f['spreadsheetId'], f['sheets'][0]['properties']['sheetId']

    # --- nội dung ---
    vals = [[L['title']], [L['info']], [L['guide']], L['header']]
    for it in L['rows']:
        vals.append([it['label']] if it['kind'] == 'group' else [it['stt'], it['khoa'], '', ''])
    _chay('ghi nội dung vào bảng', sheets.spreadsheets().values().update(
        spreadsheetId=ss_id, range="'%s'!A1" % SHEET_TITLE, valueInputOption='RAW', body={'values': vals}))

    # --- định dạng ---
    rq = []
    for i, w in enumerate((60, 70, 160, 300)):
        rq.append({'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': i, 'endIndex': i + 1},
            'properties': {'pixelSize': w}, 'fields': 'pixelSize'}})
    for r in (1, 2, 3):   # dòng 1-3 gộp ô
        rq.append({'mergeCells': {'range': _range(sid, r - 1, r), 'mergeType': 'MERGE_ALL'}})
    for r, h in ((1, 44), (2, 22 * L['info_lines'] + 10), (3, 52), (4, 28)):
        rq.append({'updateDimensionProperties': {
            'range': {'sheetId': sid, 'dimension': 'ROWS', 'startIndex': r - 1, 'endIndex': r},
            'properties': {'pixelSize': h}, 'fields': 'pixelSize'}})
    rq += [
        _fmt(sid, 0, 1, 0, 4, backgroundColor=_mau(NAVY), horizontalAlignment='CENTER', verticalAlignment='MIDDLE',
             wrapStrategy='WRAP', textFormat={'bold': True, 'fontSize': 14, 'foregroundColor': _mau(WHITE)}),
        _fmt(sid, 1, 2, 0, 4, backgroundColor=_mau((0.93, 0.95, 0.98)), verticalAlignment='MIDDLE', wrapStrategy='WRAP'),
        _fmt(sid, 2, 3, 0, 4, backgroundColor=_mau((1, 0.97, 0.88)), verticalAlignment='MIDDLE', wrapStrategy='WRAP',
             textFormat={'italic': True, 'fontSize': 10, 'foregroundColor': _mau((0.48, 0.29, 0))}),
        _fmt(sid, 3, 4, 0, 4, backgroundColor=_mau(BLUE), horizontalAlignment='CENTER', verticalAlignment='MIDDLE',
             textFormat={'bold': True, 'foregroundColor': _mau(WHITE)}),
        _fmt(sid, 4, L['last'], 0, 2, horizontalAlignment='CENTER'),
        _fmt(sid, 4, L['last'], 2, 4, numberFormat={'type': 'TEXT'}),   # MSSV dạng chữ, không mất số 0
    ]
    for it in L['rows']:
        if it['kind'] == 'group':
            r = it['row'] - 1
            rq.append({'mergeCells': {'range': _range(sid, r, r + 1), 'mergeType': 'MERGE_ALL'}})
            rq.append(_fmt(sid, r, r + 1, 0, 4, backgroundColor=_mau(PALE),
                           textFormat={'bold': True, 'foregroundColor': _mau(NAVY)}, verticalAlignment='MIDDLE'))
    rq.append({'updateBorders': {'range': _range(sid, 3, L['last']), **{k: {'style': 'SOLID', 'color': _mau((0.8, 0.84, 0.88))}
               for k in ('top', 'bottom', 'left', 'right', 'innerHorizontal', 'innerVertical')}}})
    # MSSV trùng -> tô đỏ
    a, z = L['first'], L['last']
    rq.append({'addConditionalFormatRule': {'index': 0, 'rule': {
        'ranges': [_range(sid, a - 1, z, 2, 3)],
        'booleanRule': {'condition': {'type': 'CUSTOM_FORMULA', 'values': [{
            'userEnteredValue': '=AND(LEN($C%d)>0,COUNTIF($C$%d:$C$%d,$C%d)>1)' % (a, a, z, a)}]},
            'format': {'backgroundColor': _mau((0.996, 0.79, 0.79)),
                       'textFormat': {'bold': True, 'foregroundColor': _mau((0.73, 0.11, 0.11))}}}}}})
    # --- khóa cả trang, chỉ mở MSSV + Họ tên của các dòng chỉ tiêu ---
    mo = [_range(sid, r1 - 1, r2, 2, 4) for r1, r2 in _khoi_slot(L['rows'])]
    rq.append({'addProtectedRange': {'protectedRange': {
        'range': {'sheetId': sid}, 'unprotectedRanges': mo,
        'description': 'Chỉ sửa cột MSSV và Họ tên', 'warningOnly': False}}})
    _chay('định dạng và khóa ô', sheets.spreadsheets().batchUpdate(spreadsheetId=ss_id, body={'requests': rq}))

    # --- chia sẻ ---
    role = os.environ.get('TB_ANYONE_ROLE', 'writer')
    try:
        drive.permissions().create(fileId=ss_id, body={'type': 'anyone', 'role': role}, fields='id').execute()
    except Exception as e:   # bảng đã tạo xong, chỉ thiếu bước mở quyền -> vẫn trả link kèm cảnh báo
        canh_bao.append('Bảng đã tạo nhưng KHÔNG mở được quyền "ai có link đều sửa" (%s). Sinh viên sẽ chưa điền được: '
                        'hãy mở bảng, bấm Chia sẻ và chọn "Bất kỳ ai có đường liên kết" – Người chỉnh sửa.'
                        % str(getattr(e, 'reason', '') or e)[:120])
    chu = os.environ.get('TB_OWNER_EMAIL', '').strip()
    if chu and not link_co_san:
        try:
            drive.permissions().create(fileId=ss_id, body={'type': 'user', 'role': 'writer', 'emailAddress': chu},
                                       sendNotificationEmail=False, fields='id').execute()
        except Exception as e:
            canh_bao.append('Không chia sẻ được bảng cho email %s (%s).' % (chu, str(getattr(e, 'reason', '') or e)[:120]))
    return {'id': ss_id, 'link': 'https://docs.google.com/spreadsheets/d/%s/edit?usp=sharing' % ss_id,
            'chia_se_chu': chu, 'canh_bao': canh_bao}


def doc_da_dang_ky(ss_id):
    """Đọc lại bảng sinh viên đã điền -> [{'khoa','mssv','ho_ten'}] (đưa vào tb_tep.tao_xlsx_da_dang_ky)."""
    sheets, _ = _dich_vu()
    res = sheets.spreadsheets().values().get(spreadsheetId=ss_id, range="'%s'!A5:D" % SHEET_TITLE).execute()
    out = []
    for row in res.get('values', []):
        row = (row + ['', '', '', ''])[:4]
        if str(row[1]).strip().isdigit() and (str(row[2]).strip() or str(row[3]).strip()):
            out.append({'khoa': str(row[1]).strip(), 'mssv': str(row[2]).strip(), 'ho_ten': str(row[3]).strip()})
    return out