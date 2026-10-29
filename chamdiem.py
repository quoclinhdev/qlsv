# -*- coding: utf-8 -*-
"""Quản lý chấm điểm / điểm danh: Môn -> Khóa -> Nhóm -> Sinh viên -> Buổi học.

Mô-đun này độc lập với csdl.py: csdl.DB kế thừa ChamDiemMixin và nối SCHEMA bên dưới vào schema chung.
"""
import re
import sqlite3
from collections import Counter
import unicodedata
from datetime import date, datetime

# Trạng thái điểm danh (mã lưu trong CSDL -> nhãn hiển thị)
TRANG_THAI = {"cm": "Có mặt", "tr": "Trễ", "cp": "Có phép", "v": "Vắng"}
KY_HIEU = {"cm": "x", "tr": "T", "cp": "P", "v": "V"}
SO_BUOI_MAC_DINH = 6
SO_BUOI_TOI_DA = 30
SO_NHOM_TOI_DA = 50
THANG_DIEM = 10
CAU_HINH_MAC_DINH = {"ty_le_dd": 50, "hs_cm": 100, "hs_tr": 50, "hs_cp": 100, "hs_v": 0, "vang_toi_da": None}

SCHEMA = """
CREATE TABLE IF NOT EXISTS mon_hoc(
    id INTEGER PRIMARY KEY AUTOINCREMENT, ma_mon TEXT NOT NULL DEFAULT '', ten_mon TEXT NOT NULL,
    so_buoi INTEGER NOT NULL DEFAULT 6,
    ty_le_dd REAL NOT NULL DEFAULT 50,      -- % điểm điểm danh trong tổng kết (phần còn lại là điểm kiểm tra)
    hs_cm REAL NOT NULL DEFAULT 100,        -- mỗi buổi Có mặt được bao nhiêu % điểm của buổi
    hs_tr REAL NOT NULL DEFAULT 50,         -- Trễ
    hs_cp REAL NOT NULL DEFAULT 100,        -- Có phép
    hs_v REAL NOT NULL DEFAULT 0,           -- Vắng
    vang_toi_da INTEGER);                   -- vắng quá số buổi này thì điểm điểm danh = 0 (NULL = không giới hạn)
CREATE UNIQUE INDEX IF NOT EXISTS ux_mon_ma ON mon_hoc(ma_mon) WHERE ma_mon<>'';
CREATE TABLE IF NOT EXISTS nhom_hoc(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    mon_id INTEGER NOT NULL REFERENCES mon_hoc(id) ON DELETE CASCADE,
    khoa_hoc TEXT NOT NULL, stt INTEGER NOT NULL, ten TEXT NOT NULL DEFAULT '',
    UNIQUE(mon_id, khoa_hoc, stt));
CREATE TABLE IF NOT EXISTS nhom_sv(
    nhom_id INTEGER NOT NULL REFERENCES nhom_hoc(id) ON DELETE CASCADE,
    mssv TEXT NOT NULL REFERENCES sinh_vien(mssv) ON DELETE CASCADE,
    PRIMARY KEY(nhom_id, mssv));
CREATE INDEX IF NOT EXISTS ix_nhom_sv_mssv ON nhom_sv(mssv);
-- Ngày của từng buổi: tự lấy ngày điểm danh đầu tiên, có thể sửa tay
CREATE TABLE IF NOT EXISTS buoi_hoc(
    nhom_id INTEGER NOT NULL REFERENCES nhom_hoc(id) ON DELETE CASCADE,
    buoi INTEGER NOT NULL, ngay TEXT NOT NULL DEFAULT '',
    PRIMARY KEY(nhom_id, buoi));
CREATE TABLE IF NOT EXISTS diem_danh(
    nhom_id INTEGER NOT NULL, mssv TEXT NOT NULL, buoi INTEGER NOT NULL, trang_thai TEXT NOT NULL,
    PRIMARY KEY(nhom_id, mssv, buoi),
    FOREIGN KEY(nhom_id, mssv) REFERENCES nhom_sv(nhom_id, mssv) ON DELETE CASCADE);
-- Điểm kiểm tra (thang 10) của sinh viên trong nhóm
CREATE TABLE IF NOT EXISTS diem_kt(
    nhom_id INTEGER NOT NULL, mssv TEXT NOT NULL, diem REAL NOT NULL,
    PRIMARY KEY(nhom_id, mssv),
    FOREIGN KEY(nhom_id, mssv) REFERENCES nhom_sv(nhom_id, mssv) ON DELETE CASCADE);
"""


def _bo_dau(s):
    s = unicodedata.normalize("NFD", str(s or "").lower().replace("đ", "d"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def _vi_key(s):
    return unicodedata.normalize("NFD", (s or "").lower().replace("đ", "d￿"))


def chuan_khoa(s):
    """'k25', 'K 25', '25' -> '25' (khớp với cột khoa_hoc của sinh viên); chuỗi không có số thì viết hoa."""
    s = str(s or "").strip()
    m = re.search(r"\d+", s)
    return str(int(m.group())) if m else s.upper()


def hien_khoa(k):
    return f"K{k}" if str(k).isdigit() else str(k)


def chuan_ngay(s):
    """'2026-10-29' (hoặc 2026-10-9) -> '2026-10-29'; rỗng -> ''; sai -> ValueError."""
    s = str(s or "").strip()
    if not s:
        return ""
    try:
        return datetime.strptime(s, "%Y-%m-%d").date().isoformat()
    except ValueError:
        raise ValueError("Ngày không hợp lệ (cần dạng năm-tháng-ngày).")


def _gon(s):
    return " ".join(str(s or "").split())


def _so_buoi(v):
    try:
        n = int(v)
    except (TypeError, ValueError):
        raise ValueError("Số buổi phải là số nguyên.")
    if not 1 <= n <= SO_BUOI_TOI_DA:
        raise ValueError(f"Số buổi phải từ 1 đến {SO_BUOI_TOI_DA}.")
    return n


def _phan_tram(v, ten):
    try:
        x = float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        raise ValueError(f"{ten} phải là số từ 0 đến 100.")
    if not 0 <= x <= 100:
        raise ValueError(f"{ten} phải từ 0 đến 100.")
    return x


def chuan_cau_hinh(hien_co, moi, so_buoi):
    """Gộp cấu hình tính điểm mới (dict, thiếu khóa thì giữ giá trị cũ) và kiểm tra hợp lệ."""
    c = {k: (hien_co or {}).get(k, CAU_HINH_MAC_DINH[k]) for k in CAU_HINH_MAC_DINH}
    moi = moi if isinstance(moi, dict) else {}
    nhan = {"ty_le_dd": "Tỷ lệ điểm danh", "hs_cm": "Hệ số Có mặt", "hs_tr": "Hệ số Trễ", "hs_cp": "Hệ số Có phép", "hs_v": "Hệ số Vắng"}
    for k, ten in nhan.items():
        if k in moi:
            c[k] = _phan_tram(moi[k], ten)
    if "vang_toi_da" in moi:
        v = moi["vang_toi_da"]
        if v is None or str(v).strip() == "":
            c["vang_toi_da"] = None
        else:
            try:
                c["vang_toi_da"] = int(v)
            except (TypeError, ValueError):
                raise ValueError("Số buổi vắng tối đa phải là số nguyên.")
            if not 0 <= c["vang_toi_da"] <= so_buoi:
                raise ValueError(f"Số buổi vắng tối đa phải từ 0 đến {so_buoi}.")
    elif c["vang_toi_da"] is not None and c["vang_toi_da"] > so_buoi:
        raise ValueError(f"Số buổi vắng tối đa ({c['vang_toi_da']}) vượt quá số buổi mới ({so_buoi}).")
    return c


def diem_diem_danh(m, mon):
    """Điểm điểm danh (thang 10) từ {buoi: trạng thái}. Chưa điểm danh buổi nào -> None.
    Mỗi buổi = 10/số buổi điểm; nhân hệ số theo trạng thái; vắng quá giới hạn -> 0. Buổi chưa điểm danh tính 0."""
    sb = mon["so_buoi"]
    h = {b: t for b, t in (m or {}).items() if 1 <= int(b) <= sb}
    if not h:
        return None
    gh = mon["vang_toi_da"]
    if gh is not None and sum(1 for t in h.values() if t == "v") > gh:
        return 0.0
    return round(THANG_DIEM * sum(mon["hs_" + t] / 100 for t in h.values()) / sb, 2)


def diem_tong_ket(dd, kt, mon):
    """Tổng kết = ĐD × tỷ lệ ĐD + KT × (100 − tỷ lệ ĐD). Thiếu thành phần có tỷ lệ > 0 -> None (chưa đủ điểm)."""
    w_dd = float(mon["ty_le_dd"]); w_kt = 100 - w_dd
    if (w_dd > 0 and dd is None) or (w_kt > 0 and kt is None):
        return None
    return round(((dd or 0) * w_dd + (kt or 0) * w_kt) / 100, 2)


class ChamDiemMixin:
    def _migrate_cham_diem(self):
        """DB cũ: thêm cột tên nhóm (rỗng = hiển thị 'Nhóm <số>')."""
        cols = [r[1] for r in self.con.execute("PRAGMA table_info(nhom_hoc)")]
        if cols and "ten" not in cols:
            self.con.execute("ALTER TABLE nhom_hoc ADD COLUMN ten TEXT NOT NULL DEFAULT ''")
            self.con.commit()
        # DB cũ: thêm cấu hình tính điểm cho môn học
        cols = [r[1] for r in self.con.execute("PRAGMA table_info(mon_hoc)")]
        if cols:
            for c, kieu in (("ty_le_dd", "REAL NOT NULL DEFAULT 50"), ("hs_cm", "REAL NOT NULL DEFAULT 100"),
                            ("hs_tr", "REAL NOT NULL DEFAULT 50"), ("hs_cp", "REAL NOT NULL DEFAULT 100"),
                            ("hs_v", "REAL NOT NULL DEFAULT 0"), ("vang_toi_da", "INTEGER")):
                if c not in cols:
                    self.con.execute(f"ALTER TABLE mon_hoc ADD COLUMN {c} {kieu}")
            self.con.commit()

    # ---------- môn học
    def ds_mon(self):
        return self.con.execute("SELECT * FROM mon_hoc ORDER BY ten_mon COLLATE NOCASE, id").fetchall()

    def lay_mon(self, mon_id):
        return self.con.execute("SELECT * FROM mon_hoc WHERE id=?", (mon_id,)).fetchone()

    def them_mon(self, ma_mon, ten_mon, so_buoi=SO_BUOI_MAC_DINH, cau_hinh=None):
        ma, ten, sb = _gon(ma_mon).upper(), _gon(ten_mon), _so_buoi(so_buoi)
        if not ten:
            raise ValueError("Tên môn không được để trống.")
        c = chuan_cau_hinh(None, cau_hinh, sb)
        try:
            cur = self.con.execute(
                "INSERT INTO mon_hoc(ma_mon,ten_mon,so_buoi,ty_le_dd,hs_cm,hs_tr,hs_cp,hs_v,vang_toi_da) VALUES(?,?,?,?,?,?,?,?,?)",
                (ma, ten, sb, c["ty_le_dd"], c["hs_cm"], c["hs_tr"], c["hs_cp"], c["hs_v"], c["vang_toi_da"]))
        except sqlite3.IntegrityError:
            raise ValueError("Mã môn đã tồn tại.")
        self.con.commit()
        return cur.lastrowid

    def sua_mon(self, mon_id, ma_mon, ten_mon, so_buoi, cau_hinh=None):
        cu = self.lay_mon(mon_id)
        if not cu:
            raise ValueError("Không tìm thấy môn.")
        ma, ten, sb = _gon(ma_mon).upper(), _gon(ten_mon), _so_buoi(so_buoi)
        c = chuan_cau_hinh(dict(cu), cau_hinh, sb)
        if not ten:
            raise ValueError("Tên môn không được để trống.")
        cao = self.con.execute(
            "SELECT COALESCE(MAX(d.buoi),0) FROM diem_danh d JOIN nhom_hoc n ON n.id=d.nhom_id WHERE n.mon_id=?",
            (mon_id,)).fetchone()[0]
        if sb < cao:
            raise ValueError(f"Đã có điểm danh đến buổi {cao}, không thể giảm số buổi xuống {sb}.")
        try:
            self.con.execute(
                "UPDATE mon_hoc SET ma_mon=?,ten_mon=?,so_buoi=?,ty_le_dd=?,hs_cm=?,hs_tr=?,hs_cp=?,hs_v=?,vang_toi_da=? WHERE id=?",
                (ma, ten, sb, c["ty_le_dd"], c["hs_cm"], c["hs_tr"], c["hs_cp"], c["hs_v"], c["vang_toi_da"], mon_id))
        except sqlite3.IntegrityError:
            raise ValueError("Mã môn đã tồn tại.")
        self.con.execute(  # bỏ ngày của các buổi vượt quá số buổi mới (chưa có điểm danh)
            "DELETE FROM buoi_hoc WHERE buoi>? AND nhom_id IN (SELECT id FROM nhom_hoc WHERE mon_id=?)", (sb, mon_id))
        self.con.commit()

    def xoa_mon(self, mon_id):
        return self._xoa("DELETE FROM mon_hoc WHERE id=?", (mon_id,))

    # ---------- khóa / nhóm
    def ds_khoa_cua_mon(self, mon_id):
        rows = self.con.execute(
            """SELECT n.khoa_hoc, COUNT(DISTINCT n.id) AS so_nhom, COUNT(ns.mssv) AS so_sv
               FROM nhom_hoc n LEFT JOIN nhom_sv ns ON ns.nhom_id=n.id WHERE n.mon_id=? GROUP BY n.khoa_hoc""",
            (mon_id,)).fetchall()
        out = [dict(r) for r in rows]
        tong = Counter(chuan_khoa(r[0]) for r in self.con.execute("SELECT khoa_hoc FROM sinh_vien") if str(r[0] or "").strip())
        da = Counter(r[0] for r in self.con.execute(
            """SELECT DISTINCT h.khoa_hoc, s.mssv, s.khoa_hoc FROM nhom_sv ns JOIN nhom_hoc h ON h.id=ns.nhom_id
               JOIN sinh_vien s ON s.mssv=ns.mssv WHERE h.mon_id=?""", (mon_id,)) if chuan_khoa(r[2]) == r[0])
        for x in out:  # chua_xep = SV của khóa này chưa thuộc nhóm nào của môn
            x["tong_sv"] = tong.get(x["khoa_hoc"], 0)
            x["chua_xep"] = max(0, x["tong_sv"] - da.get(x["khoa_hoc"], 0))
        out.sort(key=lambda x: (0, -int(x["khoa_hoc"])) if x["khoa_hoc"].isdigit() else (1, x["khoa_hoc"]))
        return out

    def ds_nhom(self, mon_id, khoa_hoc):
        return self.con.execute(
            """SELECT n.id, n.stt, n.ten, COALESCE(NULLIF(n.ten,''),'Nhóm '||n.stt) AS ten_hien, COUNT(ns.mssv) AS so_sv FROM nhom_hoc n
               LEFT JOIN nhom_sv ns ON ns.nhom_id=n.id WHERE n.mon_id=? AND n.khoa_hoc=?
               GROUP BY n.id ORDER BY n.stt""", (mon_id, khoa_hoc)).fetchall()

    def lay_nhom(self, nhom_id):
        return self.con.execute(
            """SELECT n.*, COALESCE(NULLIF(n.ten,''),'Nhóm '||n.stt) AS ten_hien, m.ten_mon, m.ma_mon, m.so_buoi, m.ty_le_dd, m.hs_cm, m.hs_tr, m.hs_cp, m.hs_v, m.vang_toi_da FROM nhom_hoc n JOIN mon_hoc m ON m.id=n.mon_id
               WHERE n.id=?""", (nhom_id,)).fetchone()

    def tao_nhom(self, mon_id, khoa_hoc, so_nhom, ten_nhom=None):
        """Đảm bảo khóa này của môn có đủ nhóm 1..so_nhom (nhóm đã có được giữ nguyên). Trả về (khóa chuẩn, số nhóm mới)."""
        if not self.lay_mon(mon_id):
            raise ValueError("Không tìm thấy môn.")
        khoa = chuan_khoa(khoa_hoc)
        if not khoa:
            raise ValueError("Hãy nhập khóa (vd K25).")
        ten_nhom = list(dict.fromkeys(_gon(t) for t in (ten_nhom or []) if _gon(t)))
        if ten_nhom:  # tên riêng (A1, A2, B1…): mỗi tên một nhóm mới, bỏ qua tên đã có trong khóa
            co = {r[0].lower() for r in self.con.execute(
                "SELECT COALESCE(NULLIF(ten,''),'Nhóm '||stt) FROM nhom_hoc WHERE mon_id=? AND khoa_hoc=?", (mon_id, khoa))}
            stt = self.con.execute("SELECT COALESCE(MAX(stt),0) FROM nhom_hoc WHERE mon_id=? AND khoa_hoc=?",
                                   (mon_id, khoa)).fetchone()[0]
            moi = 0
            for t in ten_nhom:
                if len(t) > 40:
                    raise ValueError("Tên nhóm tối đa 40 ký tự.")
                if t.lower() in co:
                    continue
                stt += 1
                moi += self.con.execute("INSERT INTO nhom_hoc(mon_id,khoa_hoc,stt,ten) VALUES(?,?,?,?)",
                                        (mon_id, khoa, stt, t)).rowcount
            self.con.commit()
            return khoa, moi
        try:
            n = int(so_nhom)
        except (TypeError, ValueError):
            raise ValueError("Số nhóm phải là số nguyên.")
        if not 1 <= n <= SO_NHOM_TOI_DA:
            raise ValueError(f"Số nhóm phải từ 1 đến {SO_NHOM_TOI_DA}.")
        moi = sum(self.con.execute("INSERT OR IGNORE INTO nhom_hoc(mon_id,khoa_hoc,stt) VALUES(?,?,?)",
                                   (mon_id, khoa, s)).rowcount for s in range(1, n + 1))
        self.con.commit()
        return khoa, moi

    def doi_ten_nhom(self, nhom_id, ten):
        """Đặt lại tên nhóm (rỗng = về mặc định 'Nhóm <số>'); không trùng tên nhóm khác cùng môn + khóa."""
        n = self.lay_nhom(nhom_id)
        if not n:
            raise ValueError("Không tìm thấy nhóm.")
        ten = _gon(ten)
        if len(ten) > 40:
            raise ValueError("Tên nhóm tối đa 40 ký tự.")
        moi = ten or f"Nhóm {n['stt']}"
        for r in self.con.execute(
                "SELECT COALESCE(NULLIF(ten,''),'Nhóm '||stt) FROM nhom_hoc WHERE mon_id=? AND khoa_hoc=? AND id<>?",
                (n["mon_id"], n["khoa_hoc"], nhom_id)):
            if r[0].lower() == moi.lower():
                raise ValueError(f"Đã có nhóm tên “{r[0]}” trong khóa này.")
        self.con.execute("UPDATE nhom_hoc SET ten=? WHERE id=?", (ten, nhom_id))
        self.con.commit()
        return moi

    def xoa_nhom(self, nhom_id):
        return self._xoa("DELETE FROM nhom_hoc WHERE id=?", (nhom_id,))

    def _xoa(self, sql, args):
        self.con.execute("PRAGMA foreign_keys=ON")
        n = self.con.execute(sql, args).rowcount
        self.con.commit()
        return n

    # ---------- sinh viên trong nhóm
    def ds_sv_nhom(self, nhom_id):
        rows = self.con.execute(
            """SELECT s.mssv,s.ho_lot,s.ten,s.lop,s.ma_lop,s.khoa_hoc FROM nhom_sv ns
               JOIN sinh_vien s ON s.mssv=ns.mssv WHERE ns.nhom_id=?""", (nhom_id,)).fetchall()
        return sorted(rows, key=lambda r: (_vi_key(r["ten"]), _vi_key(r["ho_lot"]), r["mssv"]))

    def ung_vien(self, nhom_id, tu_khoa="", tat_ca_khoa=False, gioi_han=300):
        """Sinh viên chưa thuộc nhóm nào của (môn, khóa) này. Mặc định chỉ lấy SV cùng khóa."""
        n = self.lay_nhom(nhom_id)
        if not n:
            raise ValueError("Không tìm thấy nhóm.")
        da = {r[0] for r in self.con.execute(
            """SELECT ns.mssv FROM nhom_sv ns JOIN nhom_hoc h ON h.id=ns.nhom_id
               WHERE h.mon_id=? AND h.khoa_hoc=?""", (n["mon_id"], n["khoa_hoc"]))}
        k = _bo_dau(tu_khoa)
        out = []
        for r in self.con.execute("SELECT mssv,ho_lot,ten,lop,ma_lop,khoa_hoc FROM sinh_vien"):
            if r["mssv"] in da:
                continue
            if not tat_ca_khoa and chuan_khoa(r["khoa_hoc"]) != n["khoa_hoc"]:
                continue
            if k and k not in _bo_dau(f"{r['mssv']} {r['ho_lot']} {r['ten']} {r['ma_lop']}"):
                continue
            out.append(r)
        out.sort(key=lambda r: (r["ma_lop"] or "", _vi_key(r["ten"]), _vi_key(r["ho_lot"])))
        return len(out), [dict(r) for r in out[:gioi_han]]

    def them_sv_nhom(self, nhom_id, mssvs):
        n = self.lay_nhom(nhom_id)
        if not n:
            raise ValueError("Không tìm thấy nhóm.")
        res = {"da_them": 0, "da_co": 0, "o_nhom_khac": [], "khong_co": []}
        for m in dict.fromkeys(str(x).strip() for x in mssvs if str(x).strip()):
            if not self.con.execute("SELECT 1 FROM sinh_vien WHERE mssv=?", (m,)).fetchone():
                res["khong_co"].append(m)
                continue
            r = self.con.execute(
                """SELECT h.id, COALESCE(NULLIF(h.ten,''),'Nhóm '||h.stt) AS ten_hien FROM nhom_sv ns JOIN nhom_hoc h ON h.id=ns.nhom_id
                   WHERE h.mon_id=? AND h.khoa_hoc=? AND ns.mssv=?""", (n["mon_id"], n["khoa_hoc"], m)).fetchone()
            if r:
                if r["id"] == nhom_id:
                    res["da_co"] += 1
                else:
                    res["o_nhom_khac"].append(f"{m} ({r['ten_hien']})")
                continue
            self.con.execute("INSERT INTO nhom_sv(nhom_id,mssv) VALUES(?,?)", (nhom_id, m))
            res["da_them"] += 1
        self.con.commit()
        return res

    def chuyen_sv_nhom(self, nhom_id, den_id, mssvs, kem_diem_danh=True):
        """Chuyển sinh viên sang nhóm khác của cùng môn + khóa. kem_diem_danh=True: mang theo điểm danh đã có;
        False: xóa điểm danh của SV trong nhóm cũ."""
        a, b = self.lay_nhom(nhom_id), self.lay_nhom(den_id)
        if not a or not b:
            raise ValueError("Không tìm thấy nhóm.")
        if nhom_id == den_id:
            raise ValueError("Nhóm đích trùng nhóm hiện tại.")
        if a["mon_id"] != b["mon_id"] or a["khoa_hoc"] != b["khoa_hoc"]:
            raise ValueError("Chỉ chuyển được giữa các nhóm cùng môn và cùng khóa.")
        res = {"da_chuyen": 0, "bo_qua": [], "ten_den": b["ten_hien"]}
        for m in dict.fromkeys(str(x).strip() for x in mssvs if str(x).strip()):
            if not self.con.execute("SELECT 1 FROM nhom_sv WHERE nhom_id=? AND mssv=?", (nhom_id, m)).fetchone() \
                    or self.con.execute("SELECT 1 FROM nhom_sv WHERE nhom_id=? AND mssv=?", (den_id, m)).fetchone():
                res["bo_qua"].append(m)
                continue
            self.con.execute("INSERT INTO nhom_sv(nhom_id,mssv) VALUES(?,?)", (den_id, m))
            if kem_diem_danh:
                self.con.execute("UPDATE diem_danh SET nhom_id=? WHERE nhom_id=? AND mssv=?", (den_id, nhom_id, m))
                self.con.execute("UPDATE diem_kt SET nhom_id=? WHERE nhom_id=? AND mssv=?", (den_id, nhom_id, m))
            else:
                self.con.execute("DELETE FROM diem_danh WHERE nhom_id=? AND mssv=?", (nhom_id, m))
                self.con.execute("DELETE FROM diem_kt WHERE nhom_id=? AND mssv=?", (nhom_id, m))
            self.con.execute("DELETE FROM nhom_sv WHERE nhom_id=? AND mssv=?", (nhom_id, m))
            res["da_chuyen"] += 1
        self.con.commit()
        return res

    def xoa_sv_nhom(self, nhom_id, mssvs):
        ds = [str(x).strip() for x in mssvs if str(x).strip()]
        n = 0
        for m in ds:  # xóa điểm danh trước (không phụ thuộc cấu hình khóa ngoại)
            self.con.execute("DELETE FROM diem_danh WHERE nhom_id=? AND mssv=?", (nhom_id, m))
            self.con.execute("DELETE FROM diem_kt WHERE nhom_id=? AND mssv=?", (nhom_id, m))
            n += self.con.execute("DELETE FROM nhom_sv WHERE nhom_id=? AND mssv=?", (nhom_id, m)).rowcount
        self.con.commit()
        return n

    # ---------- điểm danh
    def du_lieu_diem_danh(self, nhom_id):
        """Trả về (dd, ngay): dd = {mssv: {buoi: trạng thái}}, ngay = {buoi: 'YYYY-MM-DD'}."""
        dd = {}
        for r in self.con.execute("SELECT mssv,buoi,trang_thai FROM diem_danh WHERE nhom_id=?", (nhom_id,)):
            dd.setdefault(r[0], {})[str(r[1])] = r[2]
        ngay = {str(r[0]): r[1] for r in self.con.execute(
            "SELECT buoi,ngay FROM buoi_hoc WHERE nhom_id=? AND ngay<>''", (nhom_id,))}
        return dd, ngay

    def _kiem_buoi(self, nhom_id, buoi):
        n = self.lay_nhom(nhom_id)
        if not n:
            raise ValueError("Không tìm thấy nhóm.")
        try:
            b = int(buoi)
        except (TypeError, ValueError):
            raise ValueError("Buổi không hợp lệ.")
        if not 1 <= b <= n["so_buoi"]:
            raise ValueError(f"Buổi phải từ 1 đến {n['so_buoi']}.")
        return b

    def ngay_buoi(self, nhom_id, buoi):
        r = self.con.execute("SELECT ngay FROM buoi_hoc WHERE nhom_id=? AND buoi=?", (nhom_id, buoi)).fetchone()
        return r[0] if r else ""

    def _tu_dien_ngay(self, nhom_id, buoi):
        """Buổi chưa có ngày -> tự lấy ngày hôm nay (ngày đang điểm danh). Buổi đã có ngày thì giữ nguyên."""
        self.con.execute("INSERT OR IGNORE INTO buoi_hoc(nhom_id,buoi,ngay) VALUES(?,?,'')", (nhom_id, buoi))
        self.con.execute("UPDATE buoi_hoc SET ngay=? WHERE nhom_id=? AND buoi=? AND ngay=''",
                         (date.today().isoformat(), nhom_id, buoi))

    def luu_diem_danh(self, nhom_id, mssv, buoi, trang_thai):
        """Ghi điểm danh một sinh viên ở một buổi. trang_thai rỗng = xóa. Trả về ngày của buổi."""
        b = self._kiem_buoi(nhom_id, buoi)
        tt = (trang_thai or "").strip()
        if tt and tt not in TRANG_THAI:
            raise ValueError("Trạng thái điểm danh không hợp lệ.")
        if not self.con.execute("SELECT 1 FROM nhom_sv WHERE nhom_id=? AND mssv=?", (nhom_id, mssv)).fetchone():
            raise ValueError("Sinh viên không thuộc nhóm này.")
        if tt:
            self.con.execute(
                """INSERT INTO diem_danh(nhom_id,mssv,buoi,trang_thai) VALUES(?,?,?,?)
                   ON CONFLICT(nhom_id,mssv,buoi) DO UPDATE SET trang_thai=excluded.trang_thai""",
                (nhom_id, mssv, b, tt))
            self._tu_dien_ngay(nhom_id, b)
        else:
            self.con.execute("DELETE FROM diem_danh WHERE nhom_id=? AND mssv=? AND buoi=?", (nhom_id, mssv, b))
        self.con.commit()
        return self.ngay_buoi(nhom_id, b)

    def diem_danh_ca_buoi(self, nhom_id, buoi, trang_thai, chi_o_trong=True):
        """Đặt trạng thái cho cả nhóm ở một buổi. trang_thai rỗng = xóa điểm danh của buổi.
        Trả về ({mssv: trạng thái mới}, ngày của buổi)."""
        b = self._kiem_buoi(nhom_id, buoi)
        tt = (trang_thai or "").strip()
        if tt and tt not in TRANG_THAI:
            raise ValueError("Trạng thái điểm danh không hợp lệ.")
        ds = [r[0] for r in self.con.execute("SELECT mssv FROM nhom_sv WHERE nhom_id=?", (nhom_id,))]
        if not tt:
            self.con.execute("DELETE FROM diem_danh WHERE nhom_id=? AND buoi=?", (nhom_id, b))
            self.con.commit()
            return {m: "" for m in ds}, self.ngay_buoi(nhom_id, b)
        co = {r[0] for r in self.con.execute("SELECT mssv FROM diem_danh WHERE nhom_id=? AND buoi=?", (nhom_id, b))}
        doi = {}
        for m in ds:
            if chi_o_trong and m in co:
                continue
            self.con.execute(
                """INSERT INTO diem_danh(nhom_id,mssv,buoi,trang_thai) VALUES(?,?,?,?)
                   ON CONFLICT(nhom_id,mssv,buoi) DO UPDATE SET trang_thai=excluded.trang_thai""",
                (nhom_id, m, b, tt))
            doi[m] = tt
        if doi:
            self._tu_dien_ngay(nhom_id, b)
        self.con.commit()
        return doi, self.ngay_buoi(nhom_id, b)

    def dat_ngay_buoi(self, nhom_id, buoi, ngay):
        """Sửa tay ngày của một buổi (rỗng = bỏ ngày; lần điểm danh sau sẽ tự lấy lại ngày hôm nay)."""
        b = self._kiem_buoi(nhom_id, buoi)
        ng = chuan_ngay(ngay)
        self.con.execute("INSERT OR IGNORE INTO buoi_hoc(nhom_id,buoi,ngay) VALUES(?,?,'')", (nhom_id, b))
        self.con.execute("UPDATE buoi_hoc SET ngay=? WHERE nhom_id=? AND buoi=?", (ng, nhom_id, b))
        self.con.commit()
        return ng

    # ---------- điểm kiểm tra
    def du_lieu_diem_kt(self, nhom_id):
        """{mssv: điểm kiểm tra}."""
        return {r[0]: r[1] for r in self.con.execute("SELECT mssv,diem FROM diem_kt WHERE nhom_id=?", (nhom_id,))}

    def luu_diem_kt(self, nhom_id, mssv, diem):
        """Ghi điểm kiểm tra (thang 10) của một sinh viên; rỗng = xóa. Trả về điểm đã lưu (hoặc None)."""
        if not self.lay_nhom(nhom_id):
            raise ValueError("Không tìm thấy nhóm.")
        if not self.con.execute("SELECT 1 FROM nhom_sv WHERE nhom_id=? AND mssv=?", (nhom_id, mssv)).fetchone():
            raise ValueError("Sinh viên không thuộc nhóm này.")
        t = str(diem if diem is not None else "").strip().replace(",", ".")
        if not t:
            self.con.execute("DELETE FROM diem_kt WHERE nhom_id=? AND mssv=?", (nhom_id, mssv))
            self.con.commit()
            return None
        try:
            x = float(t)
        except ValueError:
            raise ValueError("Điểm kiểm tra phải là số.")
        if not 0 <= x <= THANG_DIEM:
            raise ValueError(f"Điểm kiểm tra phải từ 0 đến {THANG_DIEM}.")
        x = round(x, 2)
        self.con.execute(
            """INSERT INTO diem_kt(nhom_id,mssv,diem) VALUES(?,?,?)
               ON CONFLICT(nhom_id,mssv) DO UPDATE SET diem=excluded.diem""", (nhom_id, mssv, x))
        self.con.commit()
        return x