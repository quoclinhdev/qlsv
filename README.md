# Quản lý sinh viên – Flask + SQLite

## Chạy (Windows)
Bấm đúp `chay.bat` (tự tạo venv + cài thư viện), hoặc:

    pip install -r requirements.txt
    python app.py          -> http://127.0.0.1:5000

Không cần đăng nhập/mật khẩu: mở là dùng. Mặc định chỉ máy chạy mới truy cập được (127.0.0.1).
**Không mở ra mạng (HOST=0.0.0.0) nếu chưa có cách bảo vệ khác**, vì ai truy cập được cũng sửa/xóa được dữ liệu.

## Dữ liệu
- CSDL `quanly_sinhvien.db` (đổi đường dẫn bằng biến `QLSV_DB`), khóa phiên `secret.key` – đừng chia sẻ.
- Tình trạng sinh viên: tab "Tình trạng sinh viên" -> chọn học kỳ/năm học -> Nhập Excel
  (sheet XET TIEN DO, DSSV CBHT LAN 1/2). Lý do cảnh báo do admin nhập tay.
