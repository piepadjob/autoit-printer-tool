## BatchPrint v1.1.1

Natural print order release. An AutoIt tool that batch prints Word/Excel files in a folder, with page-range selection and a delay between each command sent to the printer.

### What's new in v1.1.1

- **Natural print order (numeric-aware sort)** — files are sorted numerically before printing, so numbered files print smallest → largest:
  - `1, 2, 3, ..., 9, 10, 11` — **not** `1, 10, 11, ..., 19, 2`
  - e.g. `file2.docx` now prints before `file10.docx`, `doc100.docx` prints after `doc20.docx`
  - case-insensitive, stable merge sort, correct even with leading zeros (`file007` = `file7`)

### Features

- **Print all Word/Excel files** in the selected folder (defaults to the folder containing the `.exe`)
  - Word: `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.rtf`
  - Excel: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`
- **Natural print order** — numbered files print from smallest to largest
- **Page range selection** — defaults to **page 1 only** (`1`, `2-5`, `1,3,5-8`, empty = all pages)
- **Delay (ms) between each command sent to the printer** — default `1000 ms`
- **Bilingual UI (EN/VI)**
- Option to **print all Excel sheets**
- Print log, **Stop** button, skips temporary lock files `~$...`
- A failing file is reported individually and **does not stop the whole batch**

### Requirements

- Windows
- Microsoft Word and/or Microsoft Excel installed

### Usage

1. Download `BatchPrint.exe` below
2. Put the files to print in the **same folder as `BatchPrint.exe`** (or click **Browse...** to pick another folder)
3. Enter the **pages to print** (default `1`) and the **delay (ms)** (default `1000`)
4. Click **Start printing**

### Notes

- The `.exe` is built from [`BatchPrint.au3`](../blob/main/BatchPrint.au3) with Aut2exe (x86, Unicode, max compression) — runs on both 32-bit and 64-bit Windows
- Some browsers/Windows SmartScreen may warn about an unsigned `.exe` — choose "More info" → "Run anyway" if you trust the source

---

## BatchPrint v1.1.1 (Tiếng Việt)

Bản phát hành với thứ tự in tự nhiên. Tool AutoIt in hàng loạt file Word/Excel trong một thư mục, hỗ trợ chọn khoảng trang in và delay giữa mỗi lệnh gửi tới máy in.

### Có gì mới trong v1.1.1

- **Thứ tự in tự nhiên (natural sort, nhận diện số)** — file được sắp theo số trước khi in, file đánh số in từ số nhỏ đến số lớn:
  - `1, 2, 3, ..., 9, 10, 11` — **không phải** `1, 10, 11, ..., 19, 2`
  - VD: `file2.docx` in trước `file10.docx`, `doc100.docx` in sau `doc20.docx`
  - không phân biệt hoa/thường, merge sort ổn định, đúng cả khi có số 0 đầu (`file007` = `file7`)

### Tính năng

- **In tất cả file Word/Excel** trong thư mục được chọn (mặc định là thư mục chứa file `.exe`)
  - Word: `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.rtf`
  - Excel: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`
- **Thứ tự in tự nhiên** — file đánh số in từ số nhỏ đến số lớn
- **Chọn khoảng trang in** — mặc định chỉ in **trang 1** (`1`, `2-5`, `1,3,5-8`, để trống = in toàn bộ)
- **Delay (ms) giữa mỗi lệnh gửi tới máy in** — mặc định `1000 ms`
- **Giao diện song ngữ Anh/Việt (EN/VI)**
- Tùy chọn **in toàn bộ các sheet của Excel**
- Nhật ký in, nút **Dừng**, bỏ qua file lock tạm `~$...`
- File lỗi chỉ báo riêng, **không dừng cả lô in**

### Yêu cầu

- Windows
- Đã cài **Microsoft Word** và/hoặc **Microsoft Excel**

### Cách dùng

1. Tải `BatchPrint.exe` bên dưới
2. Đặt file cần in vào **cùng thư mục với `BatchPrint.exe`** (hoặc bấm **Duyệt...** để chọn thư mục khác)
3. Nhập **trang cần in** (mặc định `1`) và **delay (ms)** (mặc định `1000`)
4. Bấm **Bắt đầu in**

---

## BatchPrint v1.1.0

Bilingual (EN/VI) release. An AutoIt tool that batch prints Word/Excel files in a folder, with page-range selection and a delay between each command sent to the printer.

### What's new in v1.1.0

- **Bilingual UI (EN/VI)** — switch language from the dropdown at the top of the window; all labels, buttons, log messages and dialogs are translated
- **English README** — [`README.md`](../blob/main/README.md) is now English, Vietnamese moved to [`README.vi.md`](../blob/main/README.vi.md), with a language switcher at the top of both

### Features

- **Print all Word/Excel files** in the selected folder (defaults to the folder containing the `.exe`)
  - Word: `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.rtf`
  - Excel: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`
- **Page range selection** — defaults to **page 1 only**
  - `1` → page 1
  - `2-5` → pages 2 to 5
  - `1,3,5-8` → pages 1, 3, and 5 through 8
  - empty → all pages
- **Delay (ms) between each command sent to the printer** — default `1000 ms`
- Option to **print all Excel sheets**
- Print log, **Stop** button, skips temporary lock files `~$...`
- A failing file is reported individually and **does not stop the whole batch**

### Requirements

- Windows
- Microsoft Word and/or Microsoft Excel installed

### Usage

1. Download `BatchPrint.exe` below
2. Put the files to print in the **same folder as `BatchPrint.exe`** (or click **Browse...** to pick another folder)
3. Enter the **pages to print** (default `1`) and the **delay (ms)** (default `1000`)
4. Click **Start printing**

### Notes

- The `.exe` is built from [`BatchPrint.au3`](../blob/main/BatchPrint.au3) with Aut2exe (x86, Unicode, max compression) — runs on both 32-bit and 64-bit Windows
- Some browsers/Windows SmartScreen may warn about an unsigned `.exe` — choose "More info" → "Run anyway" if you trust the source

---

## BatchPrint v1.1.0 (Tiếng Việt)

Bản phát hành song ngữ (EN/VI). Tool AutoIt in hàng loạt file Word/Excel trong một thư mục, hỗ trợ chọn khoảng trang in và delay giữa mỗi lệnh gửi tới máy in.

### Có gì mới trong v1.1.0

- **Giao diện song ngữ Anh/Việt (EN/VI)** — đổi ngôn ngữ ở dropdown phía trên cửa sổ; toàn bộ nhãn, nút, nhật ký và hộp thoại được dịch
- **README tiếng Anh** — [`README.md`](../blob/main/README.md) nay là tiếng Anh, bản tiếng Việt chuyển sang [`README.vi.md`](../blob/main/README.vi.md), có link chuyển ngôn ngữ ở đầu mỗi file

### Tính năng

- **In tất cả file Word/Excel** trong thư mục được chọn (mặc định là thư mục chứa file `.exe`)
- **Chọn khoảng trang in** — mặc định chỉ in **trang 1** (`1`, `2-5`, `1,3,5-8`, để trống = in toàn bộ)
- **Delay (ms) giữa mỗi lệnh gửi tới máy in** — mặc định `1000 ms`
- Tùy chọn **in toàn bộ các sheet của Excel**
- Nhật ký in, nút **Dừng**, bỏ qua file lock tạm `~$...`
- File lỗi chỉ báo riêng, **không dừng cả lô in**

### Yêu cầu

- Windows
- Đã cài **Microsoft Word** và/hoặc **Microsoft Excel**

### Cách dùng

1. Tải `BatchPrint.exe` bên dưới
2. Đặt file cần in vào **cùng thư mục với `BatchPrint.exe`** (hoặc bấm **Duyệt...** để chọn thư mục khác)
3. Nhập **trang cần in** (mặc định `1`) và **delay (ms)** (mặc định `1000`)
4. Bấm **Bắt đầu in**
