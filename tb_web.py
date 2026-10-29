# -*- coding: utf-8 -*-
"""Giao diện web: dán tin Zalo -> tự tách trường -> sửa -> tạo Google Sheet -> thông báo mới + Word + Excel.

Chạy riêng:   python tb_web.py   (http://127.0.0.1:5055)
Gắn vào FastAPI (app/main.py):
    from fastapi.middleware.wsgi import WSGIMiddleware
    from tb_web import app as tb_app
    app.mount('/thongbao', WSGIMiddleware(tb_app))
"""
import io
from datetime import date

from flask import Flask, jsonify, request, send_file

import tb_pdf
import tb_sheets
from tb_parser import parse
from tb_tep import tao_docx, tao_xlsx
from tb_van_ban import DEFAULTS, chuan_hoa, van_ban_hanh_chinh, van_ban_zalo

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024   # PDF tối đa ~15MB
MAX_PDF = 15 * 1024 * 1024
DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def _data():
    return chuan_hoa(request.get_json(silent=True) or {})


@app.get('/')
def trang_chu():
    return PAGE


@app.post('/api/phan-tich')
def api_phan_tich():
    text = (request.get_json(silent=True) or {}).get('text', '')
    kq = parse(text[:20000], date.today())
    kq.update({k: v for k, v in DEFAULTS.items()})
    kq['ngay_ban_hanh'] = date.today().isoformat()
    kq['link_sheet'] = kq['link_cu'][0] if kq['link_cu'] else ''   # link cũ chỉ để tham khảo, sẽ thay bằng sheet mới
    return jsonify(kq)


@app.post('/api/pdf')
def api_pdf():
    f = request.files.get('file')
    if not f:
        return jsonify({'loi': 'Chưa chọn file PDF.'}), 400
    data = f.read(MAX_PDF + 1)
    if len(data) > MAX_PDF:
        return jsonify({'loi': 'File quá lớn (tối đa 15MB).'}), 413
    try:
        text, cach, warn = tb_pdf.doc_pdf(data)
    except tb_pdf.LoiPDF as e:
        return jsonify({'loi': str(e)}), 422
    kq = tb_pdf.phan_tich_van_ban(text, date.today())
    kq['canh_bao'] = warn + kq['canh_bao']
    kq.update({k: v for k, v in DEFAULTS.items()})
    kq['ngay_ban_hanh'] = date.today().isoformat()
    kq['link_sheet'] = ''
    kq['van_ban'], kq['cach_doc'] = text[:60000], cach
    return jsonify(kq)


@app.post('/api/noi-dung')
def api_noi_dung():
    d = _data()
    return jsonify({'zalo': van_ban_zalo(d, date.today()), 'hanh_chinh': van_ban_hanh_chinh(d),
                    'tong': d['tong_chi_tieu']})


def _loi_google(e):
    """Rút mã lỗi + lý do + gợi ý từ lỗi Google API (HttpError) để người dùng biết cần làm gì."""
    import json as _j
    status = getattr(getattr(e, 'resp', None), 'status', '')
    msg, reason = str(e), ''
    try:
        err = _j.loads(e.content.decode('utf-8'))['error']
        msg = err.get('message', msg)
        reason = ','.join(x.get('reason', '') for x in err.get('errors', []) if x.get('reason'))
        reason = reason or err.get('status', '')
        det = ' '.join(str(x.get('reason', '')) for d in err.get('details', []) for x in [d.get('metadata', {}) or d])
        reason = (reason + ' ' + det).strip()
    except Exception:
        pass
    goi_y = ''
    low = (msg + ' ' + reason).lower()
    if 'has not been used' in low or 'service_disabled' in low or 'accessnotconfigured' in low:
        goi_y = ' → Chưa bật API: vào Google Cloud, bật cả Google Sheets API và Google Drive API cho dự án, đợi 1–2 phút rồi thử lại.'
    elif 'quota' in low or 'storage' in low:
        goi_y = ' → Hết dung lượng/quota của service account: xem hướng dẫn tôi gửi, hoặc báo lại để đổi cách tạo bảng.'
    elif status == 403 and 'sharing' in low:
        goi_y = ' → Tài khoản/Workspace đang chặn chia sẻ công khai.'
    elif status == 429:
        goi_y = ' → Gọi quá nhanh, đợi vài phút rồi thử lại.'
    return '[%s] %s%s%s' % (status or '?', msg[:300], (' (%s)' % reason) if reason else '', goi_y)


@app.post('/api/sheet')
def api_sheet():
    d = _data()
    if not d['chi_tieu']:
        return jsonify({'loi': 'Chưa có chỉ tiêu theo khóa nên chưa thể tạo bảng đăng ký.'}), 400
    link_co_san = str((request.get_json(silent=True) or {}).get('link_bang_co_san', '')).strip()[:300]
    try:
        kq = tb_sheets.tao_google_sheet(d, link_co_san)
    except tb_sheets.ChuaCauHinh as e:
        return jsonify({'loi': str(e), 'cau_hinh': True}), 503
    except tb_sheets.LoiBuoc as e:
        sa = tb_sheets.email_dich_vu()
        goi_y = (' | Email service account của bạn: %s' % sa) if sa else ''
        return jsonify({'loi': 'Google từ chối ở bước "%s": %s%s' % (e.buoc, _loi_google(e.goc), goi_y), 'sa_email': sa}), 502
    except Exception as e:   # lỗi từ Google API
        return jsonify({'loi': 'Google từ chối: ' + _loi_google(e)}), 502
    d['link_sheet'] = kq['link']
    kq['zalo'] = van_ban_zalo(d, date.today())
    kq['hanh_chinh'] = van_ban_hanh_chinh(d)
    return jsonify(kq)


@app.post('/api/docx')
def api_docx():
    return send_file(io.BytesIO(tao_docx(_data())), mimetype=DOCX, as_attachment=True, download_name='thong-bao.docx')


@app.post('/api/xlsx')
def api_xlsx():
    d = _data()
    if not d['chi_tieu']:
        return jsonify({'loi': 'Chưa có chỉ tiêu theo khóa.'}), 400
    return send_file(io.BytesIO(tao_xlsx(d)), mimetype=XLSX, as_attachment=True, download_name='bang-dang-ky.xlsx')


PAGE = r'''<!doctype html><html lang="vi"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Soạn thông báo đăng ký</title>
<style>
*{box-sizing:border-box}body{font:15px/1.45 system-ui,Segoe UI,Arial,sans-serif;margin:0;background:#f1f5f9;color:#0f172a}
header{background:#1e3a8a;color:#fff;padding:14px 20px;font-size:18px;font-weight:600}
main{max-width:1200px;margin:16px auto;padding:0 14px;display:grid;grid-template-columns:1fr 1fr;gap:14px}
@media(max-width:860px){main{grid-template-columns:1fr}}
section{background:#fff;border-radius:10px;padding:14px;box-shadow:0 1px 3px #0002}
h2{margin:0 0 10px;font-size:16px;color:#1e3a8a}
textarea,input{width:100%;font:inherit;padding:7px 9px;border:1px solid #cbd5e1;border-radius:6px}
textarea{min-height:230px;resize:vertical}
label{display:block;font-size:12.5px;color:#475569;margin:8px 0 2px}
.row{display:flex;gap:8px}.row>*{flex:1}
button{font:inherit;padding:8px 14px;border:0;border-radius:6px;background:#2563eb;color:#fff;cursor:pointer;margin:8px 6px 0 0}
button.ph{background:#e2e8f0;color:#0f172a}button.ok{background:#16a34a}
button:disabled{opacity:.5}
table{width:100%;border-collapse:collapse}td,th{padding:4px}th{font-size:12.5px;color:#475569;text-align:left}
.cb{background:#fff7e0;border:1px solid #f5d98b;color:#7a4b00;padding:8px 10px;border-radius:6px;margin-top:8px;font-size:13.5px}
.loi{background:#fee2e2;border-color:#fca5a5;color:#991b1b}
pre{white-space:pre-wrap;background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:10px;margin:0;font:14px/1.5 Consolas,monospace}
.full{grid-column:1/-1}.sm{padding:3px 9px;margin:0}
</style></head><body>
<header>📢 Soạn thông báo đăng ký – tự tạo Google Sheet, Word, Excel</header>
<main>
<section class="full"><h2>1. Dán tin nhắn Zalo</h2>
<textarea id="tho" placeholder="Dán thông báo (có biểu tượng 🗓 📍 🎁 🚌 👥 ✍️ ⏰) vào đây…"></textarea>
<button onclick="phanTich()">Phân tích tin nhắn</button><div id="cb"></div></section>

<section class="full"><h2>0. Hoặc tải file PDF văn bản (kế hoạch, thông báo…)</h2>
<input type="file" id="pdf" accept="application/pdf,.pdf" style="width:auto">
<button onclick="docPdf()" id="bp">Đọc &amp; phân tích PDF</button>
<span style="font-size:13px;color:#475569">PDF scan hoặc lỗi font tiếng Việt sẽ tự đọc bằng OCR.</span>
<div id="kp"></div>
<details id="dvb" style="display:none;margin-top:8px"><summary>Văn bản đọc được từ PDF</summary><pre id="vb" style="max-height:300px;overflow:auto"></pre></details>
<div id="tc"></div></section>

<section><h2>2. Kiểm tra / sửa thông tin</h2>
<label>Mục đích</label><input id="muc_dich">
<label>Tên sự kiện</label><input id="ten_su_kien">
<div class="row"><div><label>Ngày diễn ra</label><input type="date" id="ngay"></div>
<div><label>Từ giờ</label><input type="time" id="gio_bd"></div><div><label>Đến giờ</label><input type="time" id="gio_kt"></div></div>
<label>Thời gian ghi nguyên văn (nếu có, được dùng thay cho ngày/giờ)</label><input id="thoi_gian_text">
<label>Địa điểm</label><input id="dia_diem"><label>Quyền lợi</label><input id="quyen_loi">
<label>Phương tiện</label><input id="phuong_tien"><label>Cách đăng ký</label><input id="cach_dang_ky">
<div class="row"><div><label>Hạn chót (ngày)</label><input type="date" id="han_chot_ngay"></div>
<div><label>Hạn chót (giờ)</label><input type="time" id="han_chot_gio"></div></div>
<label>Chỉ tiêu theo khóa</label>
<table><thead><tr><th style="width:70px">Khóa</th><th style="width:80px">Số lượng</th><th>Ghi chú</th><th></th></tr></thead><tbody id="ct"></tbody></table>
<button class="ph sm" onclick="themDong()">+ Thêm khóa</button> <b id="tong"></b>
<details style="margin-top:10px"><summary>Đầu văn bản hành chính</summary>
<label>Đơn vị chủ quản</label><input id="don_vi_chu_quan"><label>Đơn vị ban hành</label><input id="don_vi">
<label>Tên đơn vị trong câu văn</label><input id="don_vi_viet">
<div class="row"><div><label>Số văn bản</label><input id="so_vb"></div><div><label>Địa danh</label><input id="dia_danh"></div>
<div><label>Ngày ban hành</label><input type="date" id="ngay_ban_hanh"></div></div>
<div class="row"><div><label>Chức danh ký</label><input id="chuc_danh_ky"></div><div><label>Người ký</label><input id="nguoi_ky"></div></div>
</details></section>

<section><h2>3. Tạo bảng đăng ký &amp; thông báo</h2>
<label>Link Google Sheet</label><input id="link_sheet" placeholder="Bấm “Tạo Google Sheet” để tự có link">
<details style="margin-top:6px"><summary style="font-size:13px">Cách dự phòng nếu Google không cho tạo bảng mới</summary>
<div style="font-size:13px;color:#475569;margin:6px 0">Tự tạo một Google Sheet trống bằng tài khoản của bạn, bấm Chia sẻ, thêm email service account làm <b>Người chỉnh sửa</b>, rồi dán link vào đây và bấm Tạo Google Sheet.</div>
<input id="link_bang_co_san" placeholder="Link bảng trống đã chia sẻ cho service account"></details>
<button class="ok" onclick="taoSheet()" id="bs">Tạo Google Sheet (cho SV chỉnh sửa)</button>
<button class="ph" onclick="capNhat()">Cập nhật thông báo</button>
<div id="kqs"></div>
<h2 style="margin-top:14px">Thông báo để đăng nhóm Zalo <button class="sm" onclick="sao('zalo')">Sao chép</button></h2>
<pre id="zalo">(chưa có)</pre>
<div style="margin-top:6px"><button class="ph" onclick="tai('docx')">⬇ Văn bản Word (đúng thể thức)</button>
<button class="ph" onclick="tai('xlsx')">⬇ Bảng đăng ký Excel</button></div>
<details style="margin-top:10px"><summary>Xem trước văn bản hành chính</summary><pre id="hc"></pre></details></section>
</main>
<script>
const F=['link_bang_co_san','thoi_gian_text','muc_dich','ten_su_kien','ngay','gio_bd','gio_kt','dia_diem','quyen_loi','phuong_tien','cach_dang_ky','han_chot_ngay','han_chot_gio','link_sheet','don_vi_chu_quan','don_vi','don_vi_viet','so_vb','dia_danh','ngay_ban_hanh','chuc_danh_ky','nguoi_ky'];
const $=i=>document.getElementById(i);
function thongBao(t,loi){$('cb').innerHTML=t?'<div class="cb'+(loi?' loi':'')+'">'+t+'</div>':''}
function themDong(k,s,g){const tr=document.createElement('tr');
 tr.innerHTML='<td><input value="'+(k||'')+'"></td><td><input type="number" min="1" value="'+(s||'')+'"></td><td><input></td><td><button class="ph sm" onclick="this.closest(\'tr\').remove();tinhTong()">✕</button></td>';
 tr.children[2].firstChild.value=g||'';tr.querySelectorAll('input').forEach(x=>x.oninput=tinhTong);$('ct').appendChild(tr);tinhTong()}
function tinhTong(){let t=0;$('ct').querySelectorAll('tr').forEach(r=>t+=parseInt(r.children[1].firstChild.value)||0);$('tong').textContent=' Tổng: '+t+' em'}
function thu(){const d={};F.forEach(k=>d[k]=$(k).value);d.chi_tieu=[...$('ct').querySelectorAll('tr')].map(r=>({khoa:r.children[0].firstChild.value,so_luong:r.children[1].firstChild.value,ghi_chu:r.children[2].firstChild.value}));return d}
async function goi(url,body){const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});return r}
async function phanTich(){const r=await goi('api/phan-tich',{text:$('tho').value});const d=await r.json();
 F.forEach(k=>{if(d[k]!==undefined)$(k).value=d[k]});$('ct').innerHTML='';(d.chi_tieu||[]).forEach(x=>themDong(x.khoa,x.so_luong,x.ghi_chu));
 if(!d.chi_tieu.length)themDong();thongBao(d.canh_bao.length?'⚠ '+d.canh_bao.join('<br>⚠ '):'✔ Đã nhận đủ các trường, hãy xem lại rồi bấm Tạo Google Sheet.',d.canh_bao.length>0);
 if(d.link_cu.length)thongBao($('cb').firstChild.innerHTML+'<br>ℹ Tin gốc có link cũ – bấm “Tạo Google Sheet” để thay bằng bảng mới.');capNhat()}
function esc(s){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
async function docPdf(){const f=$('pdf').files[0];if(!f){$('kp').innerHTML='<div class="cb loi">Hãy chọn file PDF trước.</div>';return}
 $('bp').disabled=true;$('kp').innerHTML='<div class="cb">Đang đọc PDF (bản scan có thể mất 1–2 phút)…</div>';
 const fd=new FormData();fd.append('file',f);let r,d;
 try{r=await fetch('api/pdf',{method:'POST',body:fd});d=await r.json()}catch(e){$('bp').disabled=false;$('kp').innerHTML='<div class="cb loi">Không kết nối được máy chủ.</div>';return}
 $('bp').disabled=false;if(!r.ok){$('kp').innerHTML='<div class="cb loi">'+esc(d.loi)+'</div>';return}
 F.forEach(k=>{if(d[k]!==undefined)$(k).value=d[k]});$('ct').innerHTML='';themDong();
 let h='✔ Đã đọc ('+(d.cach_doc=='ocr'?'bằng OCR':'từ chữ trong PDF')+')'+(d.so_hieu?' – văn bản số '+esc(d.so_hieu):'')+(d.loai_vb?', '+esc(d.loai_vb):'')+'.';
 if(d.canh_bao.length)h+='<br>⚠ '+d.canh_bao.map(esc).join('<br>⚠ ');
 $('kp').innerHTML='<div class="cb'+(d.canh_bao.length?'':'')+'">'+h+'<br>Kiểm tra lại các ô bên dưới rồi bấm Cập nhật thông báo.</div>';
 $('vb').textContent=d.van_ban;$('dvb').style.display='block';
 $('tc').innerHTML=d.to_chuc.length?'<details open style="margin-top:8px"><summary><b>Tổ chức thực hiện ('+d.to_chuc.length+' đơn vị)</b></summary>'+d.to_chuc.map(u=>'<p style="margin:6px 0 2px"><b>'+esc(u.don_vi)+'</b></p><ul style="margin:0 0 0 18px">'+u.nhiem_vu.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul>').join('')+'</details>':'';
 capNhat()}
async function capNhat(){const r=await goi('api/noi-dung',thu());const d=await r.json();$('zalo').textContent=d.zalo;$('hc').textContent=d.hanh_chinh}
async function taoSheet(){$('bs').disabled=true;$('kqs').innerHTML='<div class="cb">Đang tạo bảng…</div>';
 const r=await goi('api/sheet',thu());const d=await r.json();$('bs').disabled=false;
 if(!r.ok){$('kqs').innerHTML='<div class="cb loi">'+esc(d.loi)+'</div>';return}
 $('link_sheet').value=d.link;$('zalo').textContent=d.zalo;$('hc').textContent=d.hanh_chinh;
 $('kqs').innerHTML='<div class="cb">✔ Đã tạo. <a href="'+d.link+'" target="_blank">Mở bảng</a> – ai có link đều điền được vào dòng trống của mình.</div>'+(d.canh_bao||[]).map(x=>'<div class="cb loi">⚠ '+esc(x)+'</div>').join('')}
async function tai(loai){const r=await goi('api/'+loai,thu());if(!r.ok){const d=await r.json();alert(d.loi);return}
 const b=await r.blob(),a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=loai=='docx'?'thong-bao.docx':'bang-dang-ky.xlsx';a.click()}
function sao(id){navigator.clipboard.writeText($(id).textContent);}
themDong();
</script></body></html>'''

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5055, debug=False)