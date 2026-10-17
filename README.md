# BatchPrint — Batch print Word/Excel files with AutoIt

**English** | [Tiếng Việt](README.vi.md)

An AutoIt tool that prints all Word/Excel files in a specific folder (the folder in the same path as the tool), with page-range selection and a delay between each command sent to the printer.

## Download

Get the ready-to-run build (no AutoIt installation required): **[BatchPrint.exe — Releases](https://github.com/piepadjob/autoit-printer-tool/releases/latest)**

> Requirements: Windows + Microsoft Word and/or Microsoft Excel installed.

## Features

- **Print all Word/Excel files** in the selected folder (defaults to the folder containing the `.exe`)
  - Word: `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.rtf`
  - Excel: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`
- **Page range selection** — defaults to **page 1 only**
  - `1` → page 1
  - `2-5` → pages 2 to 5
  - `1,3,5-8` → pages 1, 3, and 5 through 8
  - empty → all pages
- **Delay (ms) between each command sent to the printer** — default `1000 ms`
  - Word: one print command per file
  - Excel: one print command per page range (Excel has no `Pages` argument, so ranges are sent as multiple commands)
- **Bilingual UI (EN/VI)** — switch language from the dropdown at the top of the window
- Option to **print all Excel sheets** (enabled by default) so page ranges count across the whole file
- Print log, **Stop** button, skips temporary lock files `~$...`
- A failing file (corrupted, password-protected, locked) is reported individually and **does not stop the whole batch**
- Word/Excel run hidden and are closed automatically when done

## Requirements

- Windows
- Microsoft Word and/or Microsoft Excel installed
- [AutoIt 3.3.16+](https://www.autoitscript.com/site/) (with SciTE) if you want to compile the `.exe` yourself

## Build the `.exe`

1. Install AutoIt (with SciTE for AutoIt)
2. Open [`BatchPrint.au3`](BatchPrint.au3) in SciTE
3. `Tools` → `Compile` (or press **F7**)
4. You get `BatchPrint.exe` in the same folder

> The `.au3` file is saved as **UTF-8 with BOM** so Vietnamese text displays correctly in SciTE/Windows.

## Usage

1. Put the files to print in the **same folder as the tool** (the tool defaults to the folder containing the `.exe`), or click **Browse...** to pick another folder
2. Enter the **pages to print** (default `1`)
3. Enter the **delay (ms)** between print commands (default `1000`)
4. Choose the file types (Word/Excel) and the "print all Excel sheets" option
5. Click **Start printing** — watch the log; click **Stop** to stop early

## Notes

- Avoid closing the tool while printing, to prevent hidden Word/Excel processes from hanging (if it happens, end `WINWORD.EXE` / `EXCEL.EXE` in Task Manager)
- The delay is applied **after each command sent to the printer**; increase it if the printer receives commands slowly or jams between files
- For Excel with "print all sheets" enabled, the From/To of the print command counts across the **whole file** (like Word), not per sheet
