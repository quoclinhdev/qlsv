# -*- coding: utf-8 -*-
"""Route cho module Quản lý chấm điểm (điểm danh theo môn – khóa – nhóm). Đăng ký từ app.py."""
import re
from functools import wraps

from flask import abort, jsonify, render_template, request

import chamdiem as cd


def _g(x):
    return f"{x:g}"


def _ngay_vn(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}" if iso else ""


def dang_ky(app, db, login_required, gui_excel):
    def loi(fn):  # ValueError nghiệp vụ -> JSON 400 (giao diện hiện toast)
        @wraps(fn)
        def w(*a, **k):
            try:
                return fn(*a, **k)
            except ValueError as e:
                return jsonify(error=str(e)), 400
        return w

    def _j():
        j = request.get_json(silent=True)
        return j if isinstance(j, dict) else {}

    def _int(v, ten):
        try:
            return int(v)
        except (TypeError, ValueError):
            raise ValueError(f"{ten} không hợp lệ.")

    def _ds(j):
        ds = j.get("mssvs")
        if not isinstance(ds, list) or not ds:
            raise ValueError("Chưa chọn sinh viên nào.")
        return ds

    # ------------------------------------------------------------ trang chính
    @app.get("/cham-diem")
    @login_required
    def cham_diem():
        d = db()
        mons = d.ds_mon()
        try:
            mid = int(request.args.get("mon", ""))
        except ValueError:
            mid = None
        mon = next((m for m in mons if m["id"] == mid), mons[0] if mons else None)
        khoas, nhoms, nhom, kh, svs, du_lieu, tk = [], [], None, "", [], {}, {}
        if mon:
            khoas = d.ds_khoa_cua_mon(mon["id"])
            k0 = cd.chuan_khoa(request.args.get("kh", ""))
            kh = k0 if any(k["khoa_hoc"] == k0 for k in khoas) else (khoas[0]["khoa_hoc"] if khoas else "")
            if kh:
                tk = next(k for k in khoas if k["khoa_hoc"] == kh)
                nhoms = d.ds_nhom(mon["id"], kh)
                try:
                    nid = int(request.args.get("nhom", ""))
                except ValueError:
                    nid = None
                nhom = next((n for n in nhoms if n["id"] == nid), nhoms[0] if nhoms else None)
        if nhom:
            svs = d.ds_sv_nhom(nhom["id"])
            dd, ngay = d.du_lieu_diem_danh(nhom["id"])
            du_lieu = dict(nhom_id=nhom["id"], so_buoi=mon["so_buoi"], dd=dd, ngay=ngay, kt=d.du_lieu_diem_kt(nhom["id"]))
        return render_template("cham_diem.html", mons=mons, mon=dict(mon) if mon else None, khoas=khoas, kh=kh,
                               nhoms=[dict(n) for n in nhoms], nhom=dict(nhom) if nhom else None, tk=tk, svs=svs, du_lieu=du_lieu,
                               TRANG_THAI=cd.TRANG_THAI, KY_HIEU=cd.KY_HIEU, hien_khoa=cd.hien_khoa)

    # ------------------------------------------------------------ môn / nhóm
    @app.post("/api/cd/mon")
    @login_required
    @loi
    def api_cd_mon():
        j = _j()
        if j.get("id"):
            mid = _int(j["id"], "Môn")
            db().sua_mon(mid, j.get("ma_mon"), j.get("ten_mon"), j.get("so_buoi"), j.get("cau_hinh"))
        else:
            mid = db().them_mon(j.get("ma_mon"), j.get("ten_mon"), j.get("so_buoi", cd.SO_BUOI_MAC_DINH), j.get("cau_hinh"))
        return jsonify(id=mid)

    @app.post("/api/cd/mon/xoa")
    @login_required
    @loi
    def api_cd_mon_xoa():
        n = db().xoa_mon(_int(_j().get("id"), "Môn"))
        return jsonify(da_xoa=n)

    @app.post("/api/cd/nhom")
    @login_required
    @loi
    def api_cd_nhom():
        j = _j()
        khoa, moi = db().tao_nhom(_int(j.get("mon_id"), "Môn"), j.get("khoa_hoc"), j.get("so_nhom"),
                              re.split(r"[,;\n]+", str(j.get("ten_nhom") or "")))
        return jsonify(khoa_hoc=khoa, nhom_moi=moi)

    @app.post("/api/cd/nhom/doi-ten")
    @login_required
    @loi
    def api_cd_nhom_doi_ten():
        j = _j()
        return jsonify(ten=db().doi_ten_nhom(_int(j.get("nhom_id"), "Nhóm"), j.get("ten")))

    @app.post("/api/cd/nhom/chuyen")
    @login_required
    @loi
    def api_cd_nhom_chuyen():
        j = _j()
        return jsonify(db().chuyen_sv_nhom(_int(j.get("nhom_id"), "Nhóm"), _int(j.get("den_nhom_id"), "Nhóm đích"),
                                           _ds(j), j.get("kem_diem_danh", True) is not False))

    @app.post("/api/cd/nhom/xoa")
    @login_required
    @loi
    def api_cd_nhom_xoa():
        n = db().xoa_nhom(_int(_j().get("nhom_id"), "Nhóm"))
        return jsonify(da_xoa=n)

    # ------------------------------------------------------------ sinh viên trong nhóm
    @app.get("/api/cd/ung-vien")
    @login_required
    @loi
    def api_cd_ung_vien():
        tong, ds = db().ung_vien(_int(request.args.get("nhom_id"), "Nhóm"), request.args.get("q", ""),
                                 request.args.get("all") == "1")
        return jsonify(tong=tong, sinh_vien=ds)

    @app.post("/api/cd/nhom/them-sv")
    @login_required
    @loi
    def api_cd_them_sv():
        j = _j()
        return jsonify(db().them_sv_nhom(_int(j.get("nhom_id"), "Nhóm"), _ds(j)))

    @app.post("/api/cd/nhom/xoa-sv")
    @login_required
    @loi
    def api_cd_xoa_sv():
        j = _j()
        return jsonify(da_xoa=db().xoa_sv_nhom(_int(j.get("nhom_id"), "Nhóm"), _ds(j)))

    # ------------------------------------------------------------ điểm danh
    @app.post("/api/cd/diem-danh")
    @login_required
    @loi
    def api_cd_diem_danh():
        j = _j()
        ngay = db().luu_diem_danh(_int(j.get("nhom_id"), "Nhóm"), str(j.get("mssv", "")).strip(),
                                  j.get("buoi"), j.get("trang_thai"))
        return jsonify(ngay=ngay)

    @app.post("/api/cd/diem-danh-buoi")
    @login_required
    @loi
    def api_cd_diem_danh_buoi():
        j = _j()
        doi, ngay = db().diem_danh_ca_buoi(_int(j.get("nhom_id"), "Nhóm"), j.get("buoi"), j.get("trang_thai"),
                                           j.get("chi_o_trong", True) is not False)
        return jsonify(cap_nhat=doi, ngay=ngay)

    # ------------------------------------------------------------ điểm kiểm tra
    @app.post("/api/cd/diem-kt")
    @login_required
    @loi
    def api_cd_diem_kt():
        j = _j()
        return jsonify(diem=db().luu_diem_kt(_int(j.get("nhom_id"), "Nhóm"), str(j.get("mssv", "")).strip(), j.get("diem")))

    @app.post("/api/cd/ngay")
    @login_required
    @loi
    def api_cd_ngay():
        j = _j()
        return jsonify(ngay=db().dat_ngay_buoi(_int(j.get("nhom_id"), "Nhóm"), j.get("buoi"), j.get("ngay")))

    # ------------------------------------------------------------ xuất Excel
    @app.get("/cham-diem/xuat")
    @login_required
    def cham_diem_xuat():
        d = db()
        try:
            nhom = d.lay_nhom(int(request.args.get("nhom", "")))
        except ValueError:
            nhom = None
        if not nhom:
            abort(404)
        sb = nhom["so_buoi"]
        svs = d.ds_sv_nhom(nhom["id"])
        dd, ngay = d.du_lieu_diem_danh(nhom["id"])
        kt = d.du_lieu_diem_kt(nhom["id"])
        dien_ra = {b for m in dd.values() for b in m}
        rows = []
        for i, s in enumerate(svs, 1):
            m = dd.get(s["mssv"], {})
            co = sum(1 for t in m.values() if t in ("cm", "tr"))
            d_dd = cd.diem_diem_danh(m, nhom)
            d_kt = kt.get(s["mssv"])
            d_tk = cd.diem_tong_ket(d_dd, d_kt, nhom)
            rows.append([i, s["mssv"], s["ho_lot"], s["ten"], s["ma_lop"] or "",
                         *[cd.KY_HIEU.get(m.get(str(b), ""), "") for b in range(1, sb + 1)],
                         co, sum(1 for t in m.values() if t == "cp"), sum(1 for t in m.values() if t == "v"),
                         round(100 * co / len(dien_ra)) / 100 if dien_ra else "",
                         "" if d_dd is None else d_dd, "" if d_kt is None else d_kt, "" if d_tk is None else d_tk])
        heads = ["STT", "MSSV", "Họ lót", "Tên", "Mã lớp",
                 *[f"Buổi {b}\n{_ngay_vn(ngay.get(str(b), ''))}".strip() for b in range(1, sb + 1)],
                 "Có mặt", "Có phép", "Vắng", "% có mặt",
                 f"Điểm điểm danh\n({_g(nhom['ty_le_dd'])}%)", f"Điểm kiểm tra\n({_g(100 - nhom['ty_le_dd'])}%)", "Tổng kết"]
        pl = [("Ký hiệu", "Ý nghĩa")] + [(cd.KY_HIEU[k], v) for k, v in cd.TRANG_THAI.items()] + \
             [("% có mặt", "(Có mặt + Trễ) / số buổi đã điểm danh"),
              ("", ""), ("Cách tính điểm", ""),
              ("Tỷ lệ", f"Điểm danh {_g(nhom['ty_le_dd'])}% – Kiểm tra {_g(100 - nhom['ty_le_dd'])}%"),
              ("Điểm / buổi", f"10 / {sb} buổi; Có mặt {_g(nhom['hs_cm'])}%, Trễ {_g(nhom['hs_tr'])}%, "
                              f"Có phép {_g(nhom['hs_cp'])}%, Vắng {_g(nhom['hs_v'])}% điểm của buổi"),
              ("Giới hạn vắng", "Không giới hạn" if nhom["vang_toi_da"] is None
               else f"Vắng quá {nhom['vang_toi_da']} buổi thì điểm điểm danh = 0"),
              ("Lưu ý", "Buổi chưa điểm danh tính 0 điểm; thiếu điểm kiểm tra thì chưa có tổng kết")]
        ten = f"ĐIỂM DANH – {nhom['ten_mon']} – {cd.hien_khoa(nhom['khoa_hoc'])} – {nhom['ten_hien']}"
        tep = re.sub(r"[^\w\-]+", "_", f"DiemDanh_{nhom['ma_mon'] or nhom['mon_id']}_{cd.hien_khoa(nhom['khoa_hoc'])}_{nhom['ten_hien']}")
        return gui_excel(tep + ".xlsx", None, ten, heads, rows,
                         [6, 14, 24, 12, 14, *[11] * sb, 9, 9, 9, 10, 14, 14, 10], pl)