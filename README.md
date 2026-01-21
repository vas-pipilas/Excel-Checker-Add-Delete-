# Excel State-Based Auditor v2.0

A professional-grade Python utility designed to audit Excel files by comparing a **Target** file against an **Instruction** plan. This tool uses **State-Based Integrity Logic** to verify if requested changes (Adds/Deletes) are reflected in the final state of the data, regardless of row positioning.

---

## 🚀 Key Features

* **State-Based Auditing:** Uses "Pool Search" logic to verify data existence. Ideal for "Move" operations where data is moved between sections.
* **Deep Normalization:** Handles case-insensitivity, leading/trailing whitespace, and invisible Unicode non-breaking spaces (`\xa0`).
* **Structural Flexibility:** Automatically unmerges cells and applies forward-fill (`ffill`) for data continuity.
* **Dynamic Header Discovery:** Automatically finds the required `Action` column anywhere in the sheet.
* **Production Ready:** * Bypasses Excel's 31-character sheet name limit with recursive naming.
* Generates timestamped reports to prevent overwriting previous results.
* Standard User Permissions: Runs without administrative/elevated privileges.



---

## 🛠 How to Use

### Option A: Running the Executable (.exe)

1. Download `Excel_Auditor_v2.0.exe`.
2. Double-click to launch. No installation or Python environment is required.
3. **Permissions:** This tool runs under standard user rights; no admin password is needed.

### Option B: Running via Python

1. **Install Dependencies:**
```bash
pip install pandas xlsxwriter tqdm openpyxl

```


2. **Execute:**
```bash
python excel_auditor.py

```



---

## 📋 Data Requirements

For the audit to succeed, your files must follow these simple rules:

1. **The Action Column:** The Instruction file **must** contain a column named exactly `Action`.
* **ADD:** Validates that the row exists anywhere in the Target.
* **DELETE:** Validates that the row is completely missing from the Target.


2. **Header Matching:** Other column headers in the Instruction file must match the headers in the Target file (e.g., *IP Address*, *Category*, *ID*).
3. **Case & Space:** The tool is "forgiving"—it ignores differences in Uppercase/Lowercase and accidental trailing spaces.

---

## 📦 Developer: Packing the EXE

To re-build the standalone executable without requiring admin privileges, use **PyInstaller** with the following command:

```bash
pyinstaller --noconfirm --onefile --windowed --name "Excel_Auditor_v2.0" --clean "main.py"

```

> **Flag Guide:**
> * `--onefile`: Bundles everything into a single, portable file.
> * `--windowed`: Suppresses the terminal window for a cleaner GUI experience.
> * `--clean`: Clears the cache to ensure a fresh build of the state logic.
> 
> 

---

## 📝 Change Log (v2.0)

* **Refactor:** Moved from positional checking to Pool-based integrity checking.
* **Fix:** Resolved "False Negatives" on data-move operations.
* **Fix:** Added Unicode `\xa0` normalization for copy-pasted Excel data.
* **UX:** Added automated 5-second success countdown and persistent error pop-ups.

---

**Would you like me to also provide the `.gitignore` file content to ensure your `build/`, `dist/`, and local test Excels don't clutter your GitHub repo?**
