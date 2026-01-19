# 🛠️ Excel File Auditor v1.0

A high-performance Python tool designed to reconcile and audit Excel sheets based on structured instruction files. Built for engineers to eliminate manual data verification.

## 🚀 Key Features
- **Bidirectional Sheet Check:** Detects if sheets are missing or if extra "rogue" sheets were added.
- **The "Hunter" Logic:** Smart row-searching that finds data even if row counts don't match.
- **Auto-Unmerge:** Specifically designed to handle merged cells while protecting column integrity.
- **Multi-Format Reporting:** Generates a summary `.txt` log and a color-coded `.xlsx` for engineers.

## 📦 How to Use
1. **Close all Excel files** before starting.
2. Run `Excel_File_Editor.exe`.
3. Select your **Target File** (the one to be checked).
4. Select your **Instructions File** (the one containing the Add/Delete commands).
5. Review the reports generated in the same folder.

## ⚠️ Important Requirements
- The **Instructions** file must have an `Action` column.
- Column headers must match exactly between files (Case Sensitive).

## 🛠️ Built With
- Python 3.x
- Pandas (Data Processing)
- XlsxWriter (Formatted Excel Exports)
- PyInstaller (Stand-alone Executable)
