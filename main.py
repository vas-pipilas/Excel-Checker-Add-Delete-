import pandas as pd
from tkinter import filedialog, Tk
from tkinter import messagebox
from tqdm import tqdm
import time
import datetime


def user_input_files():
    """
    Opens Windows file dialogs to allow the user to select the two required Excel files.
    Returns: Two dictionaries of DataFrames (one per sheet) for Target and Instructions.
    """
    root = Tk()
    root.withdraw()  # Hide the main Tkinter window
    root.attributes("-topmost", True)  # Bring the file dialog to the front of all windows

    print("Please select the Target Excel file...")
    target_path = filedialog.askopenfilename(
        title="Select Target File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    print("Please select the Instructions Excel file...")
    instructions_path = filedialog.askopenfilename(
        title="Select Instructions File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    # If the user closes the dialog without selecting a file, exit the program gracefully
    if not target_path or not instructions_path:
        print("Selection cancelled.")
        exit()

    # Load every sheet from the Excel files into a dictionary { "SheetName": DataFrame }
    return pd.read_excel(target_path, sheet_name=None), pd.read_excel(instructions_path, sheet_name=None)


def unmerge_data(data_dict, is_instruction_file=False):
    """
    Handles unmerged/merged cells by forward-filling data.
    If it's an instruction file, it protects the 'Action' column from being filled
    to maintain the integrity of specific task triggers.
    """
    processed_dict = {}
    for sheet_name, df in data_dict.items():
        if is_instruction_file and len(df.columns) > 0:
            # Look for the 'Action' column (case-insensitive and trimmed)
            action_col = next((col for col in df.columns if str(col).strip().lower() == 'action'), None)

            if action_col:
                # Forward fill all columns except 'Action' to keep row logic consistent
                data_cols = [c for c in df.columns if c != action_col]
                df[data_cols] = df[data_cols].ffill()
            else:
                # If no Action column found, fill the whole sheet
                df = df.ffill()
        else:
            # For target files, we fill everything to ensure every row has complete data
            df = df.ffill()

        processed_dict[sheet_name] = df
    return processed_dict


def run_audit(target_dict, instructions_dict):
    """
    The core State-Machine Engine.
    It compares instructions against the target file by searching the entire
    target sheet (Pool Search) rather than relying on row numbers.
    """
    audit_results = {}
    skipped_no_action = []

    for sheet_name, df_inst in instructions_dict.items():
        # Match sheets by name (ignoring accidental leading/trailing spaces)
        clean_name = sheet_name.strip()
        target_sheet_key = next((k for k in target_dict.keys() if k.strip() == clean_name), None)

        if target_sheet_key:
            df_target = target_dict[target_sheet_key]
            df_inst = df_inst.copy()

            # Identify the 'Action' column in the current instruction sheet
            action_col = next((col for col in df_inst.columns if str(col).strip().lower() == 'action'), None)
            if action_col is None:
                skipped_no_action.append(sheet_name)
                continue

            # Find columns that exist in BOTH files to use as comparison keys
            common_cols = [col for col in df_inst.columns if col in df_target.columns and col != action_col]
            if not common_cols: continue

            df_inst['Audit_Status'] = 'Unknown'

            # --- NORMALIZATION LOGIC ---
            def normalize(df_subset):
                """
                Standardizes data for comparison:
                1. Convert everything to string.
                2. Replace Unicode non-breaking spaces (\xa0) with standard spaces.
                3. Convert to Uppercase and strip whitespace.
                4. Convert 'NAN' strings back to empty strings for a clean match.
                """
                return (df_subset.astype(str)
                        .replace(u'\xa0', u' ', regex=True)
                        .apply(lambda x: x.str.upper().str.strip())
                        .replace('NAN', ''))

            # Normalize the entire Target sheet once to create a "Search Pool"
            df_target_clean = normalize(df_target[common_cols])

            # Process each instruction row with a terminal progress bar
            pbar = tqdm(df_inst.iterrows(), total=len(df_inst), desc=f"Auditing {sheet_name[:20]}")

            for index, row in pbar:
                try:
                    # Normalize the current instruction row for comparison
                    row_df = pd.DataFrame([row[common_cols]])
                    row_to_check = normalize(row_df).iloc[0]

                    action = str(row[action_col]).strip().lower() if pd.notna(row[action_col]) else ""

                    # PERFORM THE POOL SEARCH: Check if this exact data exists anywhere in the target sheet
                    is_present_anywhere = (df_target_clean == row_to_check).all(axis=1).any()

                    # AUDIT LOGIC:
                    # 'add' requires the row to EXIST in target.
                    # 'delete' requires the row to NOT EXIST in target.
                    if action == 'add':
                        df_inst.at[
                            index, 'Audit_Status'] = 'PASS' if is_present_anywhere else 'FAIL (Not found in target)'
                    elif action == 'delete':
                        df_inst.at[
                            index, 'Audit_Status'] = 'PASS' if not is_present_anywhere else 'FAIL (Still exists in target)'

                except Exception as row_err:
                    df_inst.at[index, 'Audit_Status'] = f'ERROR: {str(row_err)}'

            audit_results[sheet_name] = (df_inst, action_col)
        else:
            print(f"⚠️ Skipping {sheet_name}")

    return audit_results, skipped_no_action


def export_to_txt(data_dict, filename):
    """Generates a raw, human-readable text dump of all audited sheets."""
    with open(filename, "w", encoding="utf-8") as f:
        for sheet_name, result_data in data_dict.items():
            df, action_name = result_data
            f.write(f"\n{'=' * 40}\nSHEET: {sheet_name} (Action Col: {action_name})\n{'=' * 40}\n")
            # Convert float numbers to clean strings (removing trailing .0)
            output = df.to_string(float_format=lambda x: f'{x:f}'.rstrip('0').rstrip('.'), index=False)
            f.write(output)
            f.write("\n\n")


def export_summary_report(audit_results, missing_sheets, extra_sheets, skipped_no_action, filename):
    """Generates the executive summary report with statistics and consistency checks."""
    with open(filename, "w", encoding="utf-8") as f:
        f.write("============================================================\n")
        f.write("                COMPLETE AUDIT SUMMARY LOG\n")
        f.write("============================================================\n\n")

        # 1. Structure Analysis: Which sheets were processed or ignored
        f.write("--- 1. SHEET STRUCTURE ANALYSIS ---\n")
        if audit_results:
            f.write("✅ Sheets with 'Action' column found and audited:\n")
            for name, (df, action_name) in audit_results.items():
                f.write(f"  - {name} (Detected Action Header: '{action_name}')\n")

        if skipped_no_action:
            f.write("\n❌ Sheets ignored (No 'Action' column found in Row 1):\n")
            for s in skipped_no_action:
                f.write(f"  - {s}\n")

        # 2. Consistency: Checking if sheets exist in one file but not the other
        f.write("\n" + "=" * 60 + "\n\n")
        f.write("--- 2. FILE CONSISTENCY CHECK ---\n")
        if missing_sheets: f.write("\n⚠️ MISSING SHEETS: " + ", ".join(missing_sheets) + "\n")
        if extra_sheets: f.write("\n⚠️ EXTRA SHEETS: " + ", ".join(extra_sheets) + "\n")
        if not missing_sheets and not extra_sheets:
            f.write("\n✅ Sheet names match perfectly between files.\n")

        # 3. Detailed Results: Task breakdown per sheet
        f.write("\n" + "=" * 60 + "\n\n")
        f.write("--- 3. DETAILED AUDIT RESULTS ---\n\n")

        total_tasks = 0
        total_pass = 0

        for sheet_name, (df, _) in audit_results.items():
            checked_rows = df[df['Audit_Status'] != 'Unknown'].copy()
            num_tasks = len(checked_rows)
            num_pass = len(checked_rows[checked_rows['Audit_Status'] == 'PASS'])

            total_tasks += num_tasks
            total_pass += num_pass

            f.write(f"SHEET: {sheet_name}\n")
            f.write(f"Stats: {num_tasks} Total Tasks | {num_pass} Pass | {num_tasks - num_pass} Fail\n")
            f.write("-" * 40 + "\n")
            f.write(checked_rows.to_string(index=False) if not checked_rows.empty else " (No actions processed)")
            f.write("\n" + "=" * 60 + "\n\n")

        # 4. Final Performance Totals
        if total_tasks > 0:
            success_rate = (total_pass / total_tasks) * 100
            f.write(f"--- 4. FINAL PERFORMANCE TOTALS ---\n")
            f.write(f"TOTAL INSTRUCTIONS PROCESSED : {total_tasks}\n")
            f.write(f"TOTAL SUCCESSFUL ACTIONS     : {total_pass}\n")
            f.write(f"TOTAL FAILED ACTIONS         : {total_tasks - total_pass}\n")
            f.write(f"AUDIT SUCCESS RATE           : {success_rate:.2f}%\n")


def export_to_excel_report(audit_results, missing_sheets, extra_sheets):
    """
    Generates the final Excel report with color-coded results.
    Includes smart sheet name truncation to prevent Excel crashes.
    """
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"AUDIT_FOR_ENGINEER_{timestamp}.xlsx"

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    workbook = writer.book

    # Sheet Discrepancy Tab
    if missing_sheets or extra_sheets:
        issues = [[s, "MISSING"] for s in missing_sheets] + [[s, "EXTRA"] for s in extra_sheets]
        pd.DataFrame(issues, columns=['Sheet', 'Type']).to_excel(writer, sheet_name='DISCREPANCIES', index=False)

    used_names = set()
    green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

    for sheet_name, (df, _) in audit_results.items():
        checked_rows = df[df['Audit_Status'] != 'Unknown'].copy()
        if checked_rows.empty: continue

        # Truncate sheet names to 31 chars and handle duplicates to satisfy Excel requirements
        base_name = sheet_name.strip()[:30]
        final_name = base_name
        counter = 1
        while final_name.lower() in used_names:
            suffix = f"_{counter}"
            final_name = base_name[:(31 - len(suffix))] + suffix
            counter += 1
        used_names.add(final_name.lower())

        checked_rows.to_excel(writer, sheet_name=final_name, index=False)
        worksheet = writer.sheets[final_name]

        # Apply conditional formatting to the 'Audit_Status' column
        status_col_idx = checked_rows.columns.get_loc("Audit_Status")
        worksheet.conditional_format(1, status_col_idx, len(checked_rows), status_col_idx,
                                     {'type': 'text', 'criteria': 'containing', 'value': 'PASS', 'format': green_fmt})
        worksheet.conditional_format(1, status_col_idx, len(checked_rows), status_col_idx,
                                     {'type': 'text', 'criteria': 'containing', 'value': 'FAIL', 'format': red_fmt})

        # Auto-adjust column widths for readability
        for i, col in enumerate(checked_rows.columns):
            column_len = max(checked_rows[col].astype(str).str.len().max(), len(col)) + 2
            worksheet.set_column(i, i, min(column_len, 50))

    writer.close()
    return filename


def check_sheet_consistency(target_dict, instructions_dict):
    """Compares the set of sheet names in both files."""
    target_sheets = set(k.strip() for k in target_dict.keys())
    instruction_sheets = set(k.strip() for k in instructions_dict.keys())
    return list(instruction_sheets - target_sheets), list(target_sheets - instruction_sheets)


def main():
    """Application Entry Point - Orchestrates the audit workflow."""
    print("=" * 60)
    print("              Excel File Auditor v1.2")
    print("=" * 60)

    try:
        # STEP 1: LOAD
        target_raw, inst_raw = user_input_files()

        # STEP 2: PRE-PROCESS
        print("\n[STEP 2/4] Cleaning data and unmerging cells...")
        target_clean = unmerge_data(target_raw, is_instruction_file=False)
        inst_clean = unmerge_data(inst_raw, is_instruction_file=True)
        missing, extra = check_sheet_consistency(target_clean, inst_clean)

        # STEP 3: AUDIT
        print("\n[STEP 3/4] Running Smart Audit (Pool Search Mode)...")
        final_audit_report, skipped = run_audit(target_clean, inst_clean)

        # STEP 4: EXPORT
        print("\n[STEP 4/4] Generating Reports...")
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        export_to_txt(final_audit_report, f"AUDIT_RAW_{ts}.txt")
        export_summary_report(final_audit_report, missing, extra, skipped, f"AUDIT_SUMMARY_{ts}.txt")
        excel_name = export_to_excel_report(final_audit_report, missing, extra)

        messagebox.showinfo("Audit Complete", f"Finished!\nReport: {excel_name}")

        print("\n" + "=" * 60 + "\nSUCCESS: Audit completed. Closing in 5s...")
        for i in range(5, 0, -1):
            print(f"{i}...", end=" ", flush=True)
            time.sleep(1)

    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: {e}")
        messagebox.showerror("Error", f"An error occurred:\n{e}")
        input("\nPress ENTER to close...")


if __name__ == "__main__":
    main()