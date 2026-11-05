# Fome web — cập nhật dữ liệu và xuất biểu mẫu RFI

Ứng dụng chạy cục bộ trên Windows; trình duyệt kết nối tới `127.0.0.1`. File
không được tải lên dịch vụ web bên ngoài. Biểu mẫu RFI được xuất với dữ liệu
liên kết đã chuyển thành giá trị tĩnh, vì vậy không cần Excel mở hay tính lại
công thức để hiện dữ liệu.

## Chạy ứng dụng

1. Giữ nguyên các thư mục `RFI 15A`, `RP 15A`, các file nguồn và công cụ cùng
   một thư mục.
2. Cài Python 3.10 trở lên trên laptop, bật tùy chọn **Add Python to PATH**.
3. Chạy `FOME_WEB.bat`. Lần đầu, launcher cài thư viện Python cần thiết nếu
   chưa có, rồi mở trang web trong trình duyệt.
4. Giữ cửa sổ dòng lệnh đang chạy trong suốt thời gian sử dụng; nhấn `Ctrl+C`
   trong cửa sổ đó để dừng dịch vụ.

## Cập nhật từng file nguồn

Chọn hồ sơ vị trí ở đầu trang, sau đó chọn **Chọn file mới** tại thẻ Controlog
hoặc Danh sách RFI / FORM. Mỗi lần chỉ tải một file `.xlsx`; có thể dùng tên
file nguồn khác nhau giữa các vị trí. Ứng dụng kiểm tra workbook, sheet và
tiêu đề cần thiết trước khi thay thế. Nếu file mới không đúng cấu trúc, file
đang sử dụng không bị thay thế.

TTC15A dùng các thư mục nguồn/mẫu hiện có. Để thêm TTC15, TTC11 hoặc vị trí
khác, nhập mã vị trí và chọn **Tạo hồ sơ**. Hồ sơ mới bắt đầu trống, không lấy
nhầm dữ liệu hay biểu mẫu của TTC15A. Tải riêng cả hai file nguồn cho từng hồ
sơ trước khi xuất. Nguồn được lưu trong
`Fome_Sites\<MÃ_VỊ_TRÍ>\Fome_Web_Data`; mỗi lần thay file, bản trước được lưu
trong thư mục `Backup`. File gốc TTC15A không bị sửa.

## Cập nhật biểu mẫu theo vị trí

Đảm bảo đã chọn đúng vị trí, chọn nhóm **RFI** hoặc **RP**, rồi tải lên file
Excel `.xlsx`/`.xlsm` có mã loại trong tên (ví dụ `RFI-CAI.xlsx` hoặc
`RP -CAI.xlsx`). Mẫu được kiểm tra trước khi lưu vào thư mục RFI/RP riêng của
vị trí. Mẫu cùng loại được thay thế và bản cũ được sao lưu trong thư mục
`Fome_Web_Data\Backup\Templates`. Loại biểu mẫu và trạng thái sẵn có hiển thị
trên trang.

## Xuất biểu mẫu RFI hoặc RP

Tìm theo số RFI hoặc nội dung, chọn một hoặc nhiều RFI rồi nhấn **Xuất biểu mẫu
đã chọn**. Chọn thẻ **RFI** hoặc **RP** trước khi chọn số. Một biểu mẫu được tải
về dạng `.xlsx` hoặc `.xlsm`; nhiều biểu mẫu được tải về trong một file `.zip`.
Các file được lưu thêm tại `Fome_Web_Output\<MÃ_VỊ_TRÍ>`. Ứng dụng tự chọn mẫu đúng thư
mục theo loại RFI ghi trong số hiệu và giữ bố cục, định dạng, macro của mẫu.

Các công thức liên kết trong mẫu được tính từ dữ liệu FORM, Controlog và Pipe
hiện tại rồi thay bằng giá trị tại thời điểm xuất. Kết quả không cần Excel
tính lại công thức để hiện dữ liệu.

## Tự động nhận diện và thay đổi ký tự vị trí (Trạm thi công)

- Ứng dụng tự động quét và nhận diện các ô chứa thông tin địa điểm / trạm thi công
  (ví dụ: `Trạm thi công TTC15A - An Giang/ Construction station TTC15A - An Giang` tại ô G9 hoặc tương đương).
- Khi chọn hồ sơ vị trí (ví dụ `TTC15`, `TTC15A`, `TTC11`,...) và xuất biểu mẫu RFI (hoặc RP),
  ứng dụng sẽ tự động thay thế mã trạm tương ứng với vị trí đó (ví dụ: `Trạm thi công TTC15 - An Giang/ Construction station TTC15 - An Giang`).
- Khi tạo mới hồ sơ vị trí hoặc tải lên biểu mẫu mẫu, các file mẫu cũng tự động được đồng bộ tên trạm theo mã vị trí đã chọn.

## Lưu ý với số đo

- Các mẫu RP-CAI, RP-DIC, RP-FTI, RP-FUR, RP-HTI và RP-SPI dùng hàm
  `RAND`/`RANDBETWEEN` làm một phần biểu mẫu đo. Khi xuất, ứng dụng giữ nguyên
  các hàm và các công thức phụ thuộc vào chúng theo mẫu sẵn; không thay chúng
  bằng số ngẫu nhiên cố định.
- Nếu cần dùng số đo đã được kiểm định, nhập/ghi nhận kết quả thực tế theo quy
  trình nghiệm thu trước khi phát hành biểu mẫu. Lưu ý các hàm random nguyên
  bản có thể tính lại khi mở hoặc tính toán lại workbook.
- Tất cả các biểu mẫu RP có sẵn trong `RP 15A` có thể chọn để xuất. Mẫu MRIR
  cập nhật số báo cáo và ngày từ FORM; các phép tính cố định của mẫu được giữ
  nguyên.
- Danh sách RP lấy số hiệu RFI từ sheet `FORM`; số hiệu, ngày và các mục đối
  chiếu được tính từ hai file dữ liệu hiện tại.

## Phạm vi và tính tương thích

- Có thể xuất các mẫu có sẵn trong `RFI 15A` và tất cả các mẫu trong `RP 15A`.
- Công thức tra cứu dữ liệu trong biểu mẫu xuất được chuyển thành giá trị để
  tránh phụ thuộc `XLOOKUP`, `FILTER` hoặc đường dẫn workbook ngoài.
- Các hàm Excel có sẵn trong mẫu RP để tạo số đo (ví dụ `RAND`,
  `RANDBETWEEN`) được giữ nguyên theo yêu cầu; các hàm chuẩn này cần ứng dụng
  bảng tính hỗ trợ công thức Excel để tính lại khi mở.
- Máy chạy ứng dụng cần Python; máy nhận file đầu ra chỉ cần trình đọc
  `.xlsx`/`.xlsm` tương thích. Macro trong file `.xlsm` cần ứng dụng hỗ trợ
  macro nếu muốn chạy macro.
