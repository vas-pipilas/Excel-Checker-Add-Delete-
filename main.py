import pandas as pd
from tkinter import filedialog, Tk
from tkinter import messagebox
from tqdm import tqdm
import time


def user_input_files():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    print("Please select the Target Excel file...")
    target_path = filedialog.askopenfilename(title="Select Target File", filetypes=[("Excel files", "*.xlsx *.xls")])
    print("Please select the Instructions Excel file...")
    instructions_path = filedialog.askopenfilename(title="Select Instructions File",
                                                   filetypes=[("Excel files", "*.xlsx *.xls")])

    if not target_path or not instructions_path:
        print("Selection cancelled.")
        exit()

    return pd.read_excel(target_path, sheet_name=None), pd.read_excel(instructions_path, sheet_name=None)


def unmerge_data(data_dict, is_instruction_file=False):
    processed_dict = {}
    for sheet_name, df in data_dict.items():
        if is_instruction_file and len(df.columns) > 0:
            last_col = df.columns[-1]
            data_cols = df.columns[:-1]
            df[data_cols] = df[data_cols].ffill()
        else:
            df = df.ffill()
        processed_dict[sheet_name] = df
    return processed_dict


def run_audit(target_dict, instructions_dict):
    audit_results = {}

    # Process each sheet
    for sheet_name, df_inst in instructions_dict.items():
        clean_name = sheet_name.strip()
        target_sheet_key = next((k for k in target_dict.keys() if k.strip() == clean_name), None)

        if target_sheet_key:
            df_target = target_dict[target_sheet_key]
            action_col = df_inst.columns[-1]
            audit_cols = list(df_inst.columns[1:-1])
            key_cols = [df_inst.columns[4], df_inst.columns[5]]

            df_inst = df_inst.copy()
            df_inst['Audit_Status'] = 'Unknown'

            # --- STABLE PROGRESS BAR ---
            # mininterval=1.0 means it only updates the screen ONCE per second
            # maxinterval=2.0 ensures it doesn't get lazy
            pbar = tqdm(df_inst.iterrows(), total=len(df_inst),
                        desc=f"Auditing {sheet_name[:20]}...",
                        unit="row",
                        mininterval=1.0,
                        maxinterval=2.0)

            for index, row in pbar:
                row_to_check = row[audit_cols]
                action = str(row[action_col]).strip().lower()

                # Comparison Logic
                match_condition = (df_target[audit_cols] == row_to_check).all(axis=1)
                exists_perfectly = match_condition.any()

                if action == 'add':
                    if exists_perfectly:
                        df_inst.at[index, 'Audit_Status'] = 'PASS'
                    else:
                        key_data = row[key_cols]
                        partial_match = (df_target[key_cols] == key_data).all(axis=1)
                        if partial_match.any():
                            found_at = df_target.index[partial_match].tolist()
                            df_inst.at[index, 'Audit_Status'] = f'FAIL (Found partial match at rows {found_at})'
                        else:
                            df_inst.at[index, 'Audit_Status'] = 'FAIL (Not found)'

                elif action == 'delete':
                    df_inst.at[index, 'Audit_Status'] = 'PASS' if not exists_perfectly else 'FAIL (Still Exists)'

            audit_results[sheet_name] = df_inst
        else:
            print(f"\n⚠️ Skipping {sheet_name}: Not found in Target.")

    return audit_results


def export_to_txt(data_dict, filename):
    with open(filename, "w", encoding="utf-8") as f:
        for sheet_name, df in data_dict.items():
            f.write(f"\n{'=' * 40}\nSHEET: {sheet_name}\n{'=' * 40}\n")
            # Force numbers to show without scientific notation
            output = df.to_string(float_format=lambda x: f'{x:f}'.rstrip('0').rstrip('.'), index=False)
            f.write(output)
            f.write("\n\n")


def export_summary_report(audit_results, missing_sheets, extra_sheets, filename="AUDIT_SUMMARY.txt"):
    with open(filename, "w", encoding="utf-8") as f:
        f.write("=== COMPLETE AUDIT LOG ===\n")

        # Report Missing Sheets
        if missing_sheets:
            f.write("\n⚠️  MISSING SHEETS (Expected but not found in Target):\n")
            for s in missing_sheets: f.write(f"  - {s}\n")

        # Report Extra Sheets
        if extra_sheets:
            f.write("\n⚠️  EXTRA SHEETS (Found in Target but not in Instructions):\n")
            for s in extra_sheets: f.write(f"  - {s}\n")

        f.write("\n" + "=" * 60 + "\n\n")

        total_tasks = 0
        total_pass = 0
        total_fail = 0

        for sheet_name, df in audit_results.items():
            checked_rows = df[df['Audit_Status'] != 'Unknown'].copy()
            num_tasks = len(checked_rows)
            num_pass = len(checked_rows[checked_rows['Audit_Status'] == 'PASS'])
            num_fail = num_tasks - num_pass

            total_tasks += num_tasks
            total_pass += num_pass
            total_fail += num_fail

            f.write(f"SHEET: {sheet_name}\n")
            f.write(f"Stats: {num_tasks} Total | {num_pass} Pass | {num_fail} Fail\n")
            f.write("-" * 40 + "\n")
            f.write(checked_rows.to_string(index=False))
            f.write("\n" + "=" * 60 + "\n\n")

        f.write("=== FINAL TOTALS ===\n")
        f.write(f"TOTAL INSTRUCTIONS PROCESSED: {total_tasks}\n")
        f.write(f"TOTAL SUCCESS: {total_pass}\n")
        f.write(f"TOTAL FAILED: {total_fail}\n")

    print(f"Summary report created: {filename}")


def export_to_excel_report(audit_results, missing_sheets, extra_sheets, filename="AUDIT_FOR_ENGINEER.xlsx"):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    workbook = writer.book

    # Create Discrepancy Tab if any sheet mismatches exist
    if missing_sheets or extra_sheets:
        issues = []
        for s in missing_sheets: issues.append([s, "MISSING IN TARGET"])
        for s in extra_sheets: issues.append([s, "EXTRA IN TARGET"])

        df_warn = pd.DataFrame(issues, columns=['Sheet Name', 'Discrepancy Type'])
        df_warn.to_excel(writer, sheet_name='SHEET_DISCREPANCIES', index=False)
        writer.sheets['SHEET_DISCREPANCIES'].set_tab_color('#FF9900')  # Orange tab

    # Global Summary Sheet
    summary_data = []
    for sheet_name, df in audit_results.items():
        checked = df[df['Audit_Status'] != 'Unknown']
        summary_data.append([sheet_name, len(checked), len(checked[checked['Audit_Status'] == 'PASS']),
                             len(checked) - len(checked[checked['Audit_Status'] == 'PASS'])])

    df_summary = pd.DataFrame(summary_data, columns=['Sheet Name', 'Total Tasks', 'Passed', 'Failed'])
    df_summary.to_excel(writer, sheet_name='OVERALL_SUMMARY', index=False)

    # Individual Sheets (with Red/Green formatting)
    green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

    for sheet_name, df in audit_results.items():
        checked_rows = df[df['Audit_Status'] != 'Unknown'].copy()
        if checked_rows.empty: continue

        short_name = sheet_name[:31]
        checked_rows.to_excel(writer, sheet_name=short_name, index=False)
        worksheet = writer.sheets[short_name]
        status_col_idx = checked_rows.columns.get_loc("Audit_Status")

        worksheet.conditional_format(1, status_col_idx, len(checked_rows), status_col_idx,
                                     {'type': 'text', 'criteria': 'containing', 'value': 'PASS',
                                      'format': green_format})
        worksheet.conditional_format(1, status_col_idx, len(checked_rows), status_col_idx,
                                     {'type': 'text', 'criteria': 'containing', 'value': 'FAIL', 'format': red_format})

        for i, col in enumerate(checked_rows.columns):
            column_len = max(checked_rows[col].astype(str).str.len().max(), len(col)) + 2
            worksheet.set_column(i, i, min(column_len, 50))

    writer.close()
    print(f"Excel Report Created: {filename}")


def check_sheet_consistency(target_dict, instructions_dict):
    target_sheets = set(k.strip() for k in target_dict.keys())
    instruction_sheets = set(k.strip() for k in instructions_dict.keys())

    # 1. In Instructions but NOT in Target (Missing)
    missing_in_target = list(instruction_sheets - target_sheets)

    # 2. In Target but NOT in Instructions (Extra/Unexpected)
    extra_in_target = list(target_sheets - instruction_sheets)

    return missing_in_target, extra_in_target


def main():
    # 1. Professional Terminal Header
    print("=" * 60)
    print("             Excel File Auditor v1.0")
    print("        Data Integrity & Consistency Engine")
    print("=" * 60)
    print("\n[STEP 1/4] Selecting files...")

    try:
        # 2. Load Files (Uses your existing Tkinter dialogs)
        target_raw, inst_raw = user_input_files()

        print("\n[STEP 2/4] Cleaning data and unmerging cells...")
        # Clean Merged Cells (Uses your ffill protection for Action column)
        target_clean = unmerge_data(target_raw, is_instruction_file=False)
        inst_clean = unmerge_data(inst_raw, is_instruction_file=True)

        # 3. Bidirectional Consistency Check (Missing vs Extra)
        missing, extra = check_sheet_consistency(target_clean, inst_clean)

        print("\n[STEP 3/4] Running Audit...")
        # 4. Audit Logic (The Hunter)
        final_audit_report = run_audit(target_clean, inst_clean)

        print("\n[STEP 4/4] Generating Reports...")
        # 5. Export All Reports
        # Note: We pass 'missing' and 'extra' to the functions that handle them
        export_to_txt(final_audit_report, "AUDIT_REPORT_RAW.txt")
        export_summary_report(final_audit_report, missing, extra, "AUDIT_SUMMARY.txt")
        export_to_excel_report(final_audit_report, missing, extra, "AUDIT_FOR_ENGINEER.xlsx")

        # 6. Final Success Notification
        print("\n" + "=" * 60)
        print("SUCCESS: Audit completed successfully.")
        print(f"Sheets Audited: {len(final_audit_report)}")
        if missing or extra:
            print(f"Discrepancies: Found {len(missing)} missing and {len(extra)} extra sheets.")
        print("=" * 60)

        messagebox.showinfo("Audit Complete",
                            "The audit has finished successfully!\n\n"
                            "Reports generated:\n"
                            "1. AUDIT_SUMMARY.txt\n"
                            "2. AUDIT_FOR_ENGINEER.xlsx\n"
                            "3. AUDIT_REPORT_RAW.txt")

    except PermissionError:
        error_msg = ("Permission Denied!\n\n"
                     "Please close 'AUDIT_FOR_ENGINEER.xlsx' if it is open and try again.")
        print(f"\n❌ ERROR: {error_msg}")
        messagebox.showerror("File Error", error_msg)

    except Exception as e:
        error_msg = f"An unexpected error occurred:\n{e}"
        print(f"\n❌ CRITICAL ERROR: {e}")
        messagebox.showerror("Unexpected Error", error_msg)

    # 7. THE PAUSE: This is vital for the .exe version
    # It prevents the terminal from closing until the user presses Enter.
    print("\nProcess finished.")
    input("Press ENTER to exit this window...")


if __name__ == "__main__":
    main()