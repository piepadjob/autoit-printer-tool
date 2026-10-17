# BatchPrint — In hàng loạt file Word/Excel bằng AutoIt

Tool AutoIt in tất cả file Word/Excel trong một thư mục cụ thể (thư mục nằm chung đường dẫn với tool), hỗ trợ chọn khoảng trang in và delay giữa mỗi lệnh gửi tới máy in.

## Tải về

Tải bản chạy sẵn (không cần cài AutoIt): **[BatchPrint.exe — Releases](https://github.com/piepadjob/autoit-printer-tool/releases/latest)**

> Yêu cầu: Windows + đã cài Microsoft Word và/hoặc Microsoft Excel.

## Tính năng

- **In tất cả file Word/Excel** trong thư mục được chọn
  - Word: `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.rtf`
  - Excel: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`
- **Chọn khoảng trang in** — mặc định chỉ in **trang 1**
  - `1` → trang 1
  - `2-5` → trang 2 đến 5
  - `1,3,5-8` → các trang 1, 3, từ 5 đến 8
  - để trống → in toàn bộ
- **Delay (ms) giữa mỗi lệnh gửi tới máy in** — mặc định `1000 ms`
  - Word: 1 lệnh in cho mỗi file
  - Excel: 1 lệnh in cho mỗi khoảng trang (Excel không có tham số `Pages` nên các khoảng được gửi thành nhiều lệnh)
- Tùy chọn **in toàn bộ các sheet của Excel** (mặc định bật) để khoảng trang tính theo toàn bộ file
- Nhật ký in (log), nút **Dừng**, bỏ qua file lock tạm `~$...`
- File lỗi (hỏng, bị mật khẩu, đang bị người khác mở) chỉ báo lỗi riêng, **không dừng cả lô in**
- Ứng dụng Word/Excel chạy ngầm (ẩn), tự động tắt sau khi in xong

## Yêu cầu

- Windows
- Đã cài **Microsoft Word** và/hoặc **Microsoft Excel**
- [AutoIt 3.3.16+](https://www.autoitscript.com/site/) (kèm SciTE) nếu muốn biên dịch thành `.exe`

## Biên dịch thành file `.exe`

1. Cài AutoIt (kèm SciTE for AutoIt)
2. Mở [`BatchPrint.au3`](BatchPrint.au3) bằng SciTE
3. `Tools` → `Compile` (hoặc bấm **F7**)
4. Được `BatchPrint.exe` ở cùng thư mục

> File `.au3` được lưu ở mã hóa **UTF-8 có BOM** để hiển thị đúng tiếng Việt trên SciTE/Windows.

## Cách dùng

1. Đặt file cần in vào **thư mục cùng đường dẫn với tool** (mặc định tool tự trỏ vào thư mục chứa `.exe`), hoặc bấm **Duyệt...** để chọn thư mục khác
2. Nhập **trang cần in** (mặc định `1`)
3. Nhập **delay (ms)** giữa các lệnh in (mặc định `1000`)
4. Chọn loại file cần in (Word/Excel) và tùy chọn in toàn bộ sheet Excel
5. Bấm **Bắt đầu in** — theo dõi nhật ký; bấm **Dừng** để dừng sớm

## Lưu ý

- Không nên tắt tool giữa lúc đang in để tránh tiến trình Word/Excel ngầm bị treo (nếu lỡ tắt, kết thúc tiến trình `WINWORD.EXE` / `EXCEL.EXE` trong Task Manager)
- Delay tính **sau mỗi lệnh gửi tới máy in**; tăng delay nếu máy in nhận lệnh chậm hoặc bị lỗi kẹt giữa các file
- Excel khi chọn "in toàn bộ sheet": From/To của lệnh in tính theo **toàn bộ file** (giống Word), không tính theo từng sheet riêng lẻ
