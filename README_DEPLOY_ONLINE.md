# HƯỚNG DẪN ĐƯA ỨNG DỤNG LÊN INTERNET (HOÀN TOÀN MIỄN PHÍ 24/7)

Ứng dụng Fome Web đã được tối ưu hóa toàn bộ để chạy trên môi trường Cloud (Internet). Bạn có thể đưa ứng dụng lên mạng chỉ với **3 bước đơn giản** và hoàn toàn miễn phí. Khi đưa lên internet:
- **Bất kỳ ai** cũng có thể truy cập qua điện thoại, máy tính bảng hoặc máy tính mà **không cần cài Python hay Office**.
- **Không phụ thuộc vào máy tính này:** Bạn có thể tắt máy tính ở nhà, ứng dụng trên internet vẫn hoạt động liên tục 24/7.

---

## CÁCH 1: ĐƯA LÊN RENDER.COM (KHUYÊN DÙNG - CỰC KỲ DỄ)

### Bước 1: Đưa mã nguồn lên GitHub (Miễn phí)
1. Đăng ký/đăng nhập tài khoản tại [https://github.com](https://github.com).
2. Tạo một repository mới (ví dụ đặt tên là `fome-pipeline`).
3. Mở PowerShell trong thư mục `c:\Users\LE TRUONG\Desktop\Fome` và chạy 3 lệnh sau:
   ```powershell
   git init
   git add .
   git commit -m "Khoi tao he thong Fome"
   git branch -M main
   git remote add origin https://github.com/<tai-khoan-cua-ban>/fome-pipeline.git
   git push -u origin main
   ```

### Bước 2: Kết nối và chạy trên Render.com
1. Đăng ký tài khoản miễn phí tại [https://render.com](https://render.com) (chọn đăng nhập bằng tài khoản GitHub).
2. Nhấn nút **New +** ở góc phải -> Chọn **Web Service**.
3. Chọn repository `fome-pipeline` vừa tạo ở Bước 1.
4. Render sẽ tự động nhận diện cấu hình:
   - **Name:** `fome-app` (tùy bạn đặt)
   - **Environment:** `Python`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python fome_web.py`
5. Chọn gói **Free** (Miễn phí $0/tháng).
6. Nhấn nút **Create Web Service**.

### Bước 3: Nhận đường link online
- Sau khoảng 1-2 phút, Render sẽ cấp cho bạn một đường link cố định có dạng:
  `https://fome-app.onrender.com`
- Bạn có thể gửi đường link này cho đồng nghiệp, sếp hoặc mở trên điện thoại/máy tính bất kỳ để sử dụng ngay mà không cần cài đặt gì thêm!

---

## CÁCH 2: CHIA SẺ TRONG MẠNG NỘI BỘ (LAN / WI-FI CÔNG TRÌNH)
Nếu các máy tính hoặc điện thoại đang dùng chung một mạng Wi-Fi tại công trường/văn phòng:
1. Xem địa chỉ IP của máy tính hiện tại bằng cách mở Command Prompt gõ `ipconfig` (ví dụ: `192.168.1.15`).
2. Khởi chạy file `1_Chay_Ung_Dung.bat` trên máy này.
3. Các máy tính hoặc điện thoại khác trong cùng Wi-Fi chỉ cần mở trình duyệt và gõ:
   `http://192.168.1.15:8765/`
   là dùng chung được ngay lập tức!
