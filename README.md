Here is a comprehensive `README.md` for your repository. It highlights the professional "State-Based" logic we developed and provides clear instructions for any engineer who might use your tool.

---

# Excel State-Based Auditor v2.0

A robust Python-based utility designed to audit Excel files by comparing a **Target** file against an **Instruction** plan. Unlike traditional row-by-row auditors, this tool uses **State-Based Integrity Logic** to verify if requested changes (Adds/Deletes) are reflected in the final state of the data, regardless of row positioning.

## 🚀 Key Features

* **State-Based Auditing:** Uses "Pool Search" logic to verify data existence. Ideal for "Move" operations where data is deleted from one section and added to another.
* **Deep Normalization:** Automatically handles case-insensitivity, leading/trailing whitespace, and invisible Unicode non-breaking spaces (`\xa0`).
* **Structural Flexibility:** Automatically unmerges cells and applies forward-fill (`ffill`) to ensure data continuity.
* **Dynamic Header Discovery:** Finds the required `Action` column anywhere in the sheet—no hard-coded column indices.
* **Production Ready:** * Handles Excel’s 31-character sheet name limit with recursive naming.
* Generates timestamped reports to prevent accidental data loss.
* Includes a GUI file picker and progress bars for large datasets.



---

## 📋 Requirements

* **Python 3.8+**
* **Dependencies:**
```bash
pip install pandas xlsxwriter tqdm openpyxl

```



---

## 🛠 How to Use

### 1. Prepare your Instruction File

Your **Instruction Excel** must contain a column named exactly `Action` (case-insensitive).

* **ADD:** The auditor checks if this row exists *anywhere* in the Target sheet.
* **DELETE:** The auditor checks if this row has been *removed* from the Target sheet.

> **Note:** All other columns in your Instruction file must match the header names in your Target file for the comparison to work.

### 2. Run the Script

```bash
python excel_auditor.py

```

### 3. Select Files

1. **Target File:** The modified Excel file you want to verify.
2. **Instruction File:** The plan containing the "Add" and "Delete" instructions.

### 4. Review the Report

The tool generates a new file: `AUDIT_FOR_ENGINEER_YYYYMMDD_HHMM.xlsx`.

* **PASS (Green):** The state matches the instruction.
* **FAIL (Red):** The data was either not found (for an ADD) or still exists (for a DELETE).

---

## 🧠 Logic: The "Pool of Truth"

Traditional auditors fail when a row moves from Index 10 to Index 50. This tool treats the Target sheet as a **Virtual State**.

1. It creates a **Normalized Pool** of all data in the Target sheet.
2. It converts all data to Uppercase and Stripped Strings to ensure that `$100`, `"100"`, and `" 100 "` are all treated as the same value.
3. It validates the **Existence** or **Absence** of the record to confirm the state transition was successful.

---

## 📝 Change Log (v2.0)

* **Refactor:** Moved from positional checking to Pool-based integrity checking.
* **Fix:** Resolved false negatives in "Move" operations.
* **Fix:** Added Unicode `\xa0` normalization for web-scraped or copy-pasted data.
* **UX:** Added automated success countdown and persistent error logging.

---

**Would you like me to add a "Troubleshooting" section to this README to explain how to handle common Excel data-type errors?**
