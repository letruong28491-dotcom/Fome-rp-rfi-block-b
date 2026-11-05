# FOME - HỆ THỐNG XUẤT BÁO CÁO RFI & RP (BLOCK B)

Dự án Đường ống Dẫn khí Lô B - Ô Môn (Block B - O Mon Gas Pipeline Project).
Ứng dụng Web tự động đối chiếu dữ liệu thi công từ Controlog và xuất biểu mẫu nghiệm thu RFI & RP chuẩn hóa.

---

## 🌟 Tính năng chính

- **Quản lý đa vị trí/trạm thi công (Sites):** Hỗ trợ TTC15, TTC15A, TTC11... và tự động nhân bản, đồng bộ tên trạm vào từng biểu mẫu.
- **Tự động đối chiếu dữ liệu:** Quét số mối hàn (Joint No.), số SHT, số REV từ file Controlog để điền chính xác vào báo cáo RFI & RP.
- **Tùy chỉnh thông số bản vẽ:** Cho phép cập nhật Drawing No., Weld Map No., tên kỹ sư trực tiếp trên giao diện web.
- **Chuẩn hóa biểu mẫu xuất:**
  - Định dạng bảng biểu, chữ đậm, nghiêng, căn lề và đường gạch khóa dòng trống tự động.
  - Chuyển đổi công thức liên kết thành giá trị tĩnh giúp mở file độc lập không phụ thuộc phần mềm Excel ngoài.
  - Hỗ trợ xuất lẻ file `.xlsx`/`.xlsm` hoặc xuất hàng loạt gói `.zip`.

---

## 🚀 Hướng dẫn triển khai Online 24/7 (Miễn phí trên Render.com)

Để ứng dụng chạy online liên tục cho mọi người truy cập qua điện thoại, máy tính bảng hoặc máy tính:

1. Đăng ký tài khoản tại [https://render.com](https://render.com) (chọn đăng nhập bằng GitHub).
2. Nhấn nút **New +** ở góc phải -> Chọn **Web Service**.
3. Chọn Repository **`Fome-rp-rfi-block-b`**.
4. Cấu hình triển khai:
   - **Name:** `fome-rp-rfi-block-b`
   - **Runtime:** `Python` hoặc `Docker` (đã có sẵn `Dockerfile`)
   - Nếu chọn Python:
     - **Build Command:** `pip install -r requirements.txt`
     - **Start Command:** `python fome_web.py`
   - **Plan:** Free ($0/tháng).
5. Nhấn **Create Web Service**. Sau 1-2 phút bạn sẽ nhận được đường link web hoạt động 24/7.

---

## 💻 Hướng dẫn chạy cục bộ (Local)

1. Cài đặt Python 3.10 trở lên.
2. Cài đặt thư viện phụ thuộc:
   ```bash
   pip install -r requirements.txt
   ```
3. Khởi chạy ứng dụng:
   ```bash
   python fome_web.py
   ```
4. Mở trình duyệt và truy cập: `http://localhost:8765`
