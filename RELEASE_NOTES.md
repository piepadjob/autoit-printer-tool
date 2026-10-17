## BatchPrint v1.0.0

Tool AutoIt in hàng loạt file Word/Excel trong một thư mục, hỗ trợ chọn khoảng trang in và delay giữa mỗi lệnh gửi tới máy in.

### Tính năng

- **In tất cả file Word/Excel** trong thư mục được chọn (mặc định là thư mục chứa file `.exe`)
  - Word: `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.rtf`
  - Excel: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`
- **Chọn khoảng trang in** — mặc định chỉ in **trang 1**
  - `1` → trang 1
  - `2-5` → trang 2 đến 5
  - `1,3,5-8` → các trang 1, 3, từ 5 đến 8
  - để trống → in toàn bộ
- **Delay (ms) giữa mỗi lệnh gửi tới máy in** — mặc định `1000 ms`
- Tùy chọn **in toàn bộ các sheet của Excel**
- Nhật ký in, nút **Dừng**, bỏ qua file lock tạm `~$...`
- File lỗi (hỏng, mật khẩu, đang bị mở) chỉ báo lỗi riêng, **không dừng cả lô in**

### Yêu cầu

- Windows
- Đã cài **Microsoft Word** và/hoặc **Microsoft Excel**

### Cách dùng

1. Tải `BatchPrint.exe` bên dưới
2. Đặt file cần in vào **cùng thư mục với `BatchPrint.exe`** (hoặc bấm **Duyệt...** để chọn thư mục khác)
3. Nhập **trang cần in** (mặc định `1`) và **delay (ms)** (mặc định `1000`)
4. Bấm **Bắt đầu in**

### Lưu ý

- File `.exe` được build từ [`BatchPrint.au3`](../blob/main/BatchPrint.au3) bằng Aut2exe (x86, Unicode, nén max) — chạy được trên cả Windows 32-bit và 64-bit
- Một số trình duyệt/Windows SmartScreen có thể cảnh báo file `.exe` không có chữ ký số — chọn "More info" → "Run anyway" nếu bạn tin tưởng nguồn
