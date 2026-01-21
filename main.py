import pandas as pd
from tkinter import filedialog, Tk
from tkinter import messagebox
from tqdm import tqdm
import time
import datetime


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
            # Flexible check for Action column in unmerge logic too
            action_col = next((col for col in df.columns if str(col).strip().lower() == 'action'), None)
            if action_col:
                # Fill everything EXCEPT the action column
                data_cols = [c for c in df.columns if c != action_col]
                df[data_cols] = df[data_cols].ffill()
            else:
                df = df.ffill()
        else:
            df = df.ffill()
        processed_dict[sheet_name] = df
    return processed_dict


def run_audit(target_dict, instructions_dict):
    audit_results = {}
    skipped_no_action = []

    for sheet_name, df_inst in instructions_dict.items():
        clean_name = sheet_name.strip()
        target_sheet_key = next((k for k in target_dict.keys() if k.strip() == clean_name), None)

        if target_sheet_key:
            df_target = target_dict[target_sheet_key]
            df_inst = df_inst.copy()

            action_col = next((col for col in df_inst.columns if str(col).strip().lower() == 'action'), None)
            if action_col is None:
                skipped_no_action.append(sheet_name)
                continue

            common_cols = [col for col in df_inst.columns if col in df_target.columns and col != action_col]
            if not common_cols: continue

            df_inst['Audit_Status'] = 'Unknown'

            # --- CORRECTED NORMALIZATION ---
            def normalize(df_subset):
                # We convert to string, then upper, THEN replace the 'NAN' string
                return (df_subset.astype(str)
                        .replace(u'\xa0', u' ', regex=True)
                        .apply(lambda x: x.str.upper().str.strip())
                        .replace('NAN', ''))

            # Normalize the entire Target once for the "Pool" check
            df_target_clean = normalize(df_target[common_cols])

            pbar = tqdm(df_inst.iterrows(), total=len(df_inst), desc=f"Auditing {sheet_name[:20]}")

            for index, row in pbar:
                try:
                    # Normalize the single instruction row
                    # We wrap in DataFrame because normalize expects a 2D object for .apply
                    row_df = pd.DataFrame([row[common_cols]])
                    row_to_check = normalize(row_df).iloc[0]

                    action = str(row[action_col]).strip().lower() if pd.notna(row[action_col]) else ""

                    # Check if this exact row exists anywhere in the pool
                    is_present_anywhere = (df_target_clean == row_to_check).all(axis=1).any()

                    if action == 'add':
                        if is_present_anywhere:
                            df_inst.at[index, 'Audit_Status'] = 'PASS'
                        else:
                            df_inst.at[index, 'Audit_Status'] = 'FAIL (Not found in target)'

                    elif action == 'delete':
                        if not is_present_anywhere:
                            df_inst.at[index, 'Audit_Status'] = 'PASS'
                        else:
                            df_inst.at[index, 'Audit_Status'] = 'FAIL (Still exists in target)'

                except Exception as row_err:
                    df_inst.at[index, 'Audit_Status'] = f'ERROR: {str(row_err)}'

            audit_results[sheet_name] = (df_inst, action_col)
        else:
            print(f"⚠️ Skipping {sheet_name}")

    return audit_results, skipped_no_action

def export_to_txt(data_dict, filename):
    with open(filename, "w", encoding="utf-8") as f:
        for sheet_name, result_data in data_dict.items():
            df, action_name = result_data
            f.write(f"\n{'=' * 40}\nSHEET: {sheet_name} (Action Col: {action_name})\n{'=' * 40}\n")
            output = df.to_string(float_format=lambda x: f'{x:f}'.rstrip('0').rstrip('.'), index=False)
            f.write(output)
            f.write("\n\n")


def export_summary_report(audit_results, missing_sheets, extra_sheets, skipped_no_action, filename="AUDIT_SUMMARY.txt"):
    with open(filename, "w", encoding="utf-8") as f:
        f.write("============================================================\n")
        f.write("                  COMPLETE AUDIT SUMMARY LOG\n")
        f.write("============================================================\n\n")

        # --- 1. SHEET STRUCTURE ANALYSIS ---
        f.write("--- 1. SHEET STRUCTURE ANALYSIS ---\n")
        if audit_results:
            f.write("✅ Sheets with 'Action' column found and audited:\n")
            for name, (df, action_name) in audit_results.items():
                f.write(f"  - {name} (Detected Action Header: '{action_name}')\n")

        if skipped_no_action:
            f.write("\n❌ Sheets ignored (No 'Action' column found in Row 1):\n")
            for s in skipped_no_action:
                f.write(f"  - {s}\n")

        f.write("\n" + "=" * 60 + "\n\n")

        # --- 2. FILE CONSISTENCY ---
        f.write("--- 2. FILE CONSISTENCY CHECK ---\n")
        if missing_sheets:
            f.write("\n⚠️ MISSING SHEETS: " + ", ".join(missing_sheets) + "\n")
        if extra_sheets:
            f.write("\n⚠️ EXTRA SHEETS: " + ", ".join(extra_sheets) + "\n")
        if not missing_sheets and not extra_sheets:
            f.write("\n✅ Sheet names match perfectly between files.\n")

        f.write("\n" + "=" * 60 + "\n\n")

        # --- 3. DETAILED TASK BREAKDOWN (The part we missed!) ---
        f.write("--- 3. DETAILED AUDIT RESULTS ---\n\n")

        total_tasks = 0
        total_pass = 0

        for sheet_name, (df, action_name) in audit_results.items():
            # Filter rows that were actually audited
            checked_rows = df[df['Audit_Status'] != 'Unknown'].copy()
            num_tasks = len(checked_rows)
            num_pass = len(checked_rows[checked_rows['Audit_Status'] == 'PASS'])

            total_tasks += num_tasks
            total_pass += num_pass

            f.write(f"SHEET: {sheet_name}\n")
            f.write(f"Stats: {num_tasks} Total Tasks | {num_pass} Pass | {num_tasks - num_pass} Fail\n")
            f.write("-" * 40 + "\n")

            if not checked_rows.empty:
                # This line restores the detailed table in your TXT report
                f.write(checked_rows.to_string(index=False))
            else:
                f.write(" (No actions processed in this sheet)")

            f.write("\n" + "=" * 60 + "\n\n")

        # --- 4. FINAL TOTALS ---
        f.write("--- 4. FINAL PERFORMANCE TOTALS ---\n")
        f.write(f"TOTAL INSTRUCTIONS PROCESSED : {total_tasks}\n")
        f.write(f"TOTAL SUCCESSFUL ACTIONS     : {total_pass}\n")
        f.write(f"TOTAL FAILED ACTIONS         : {total_tasks - total_pass}\n")

        if total_tasks > 0:
            success_rate = (total_pass / total_tasks) * 100
            f.write(f"AUDIT SUCCESS RATE           : {success_rate:.2f}%\n")

    print(f"Summary report created: {filename}")

def export_to_excel_report(audit_results, missing_sheets, extra_sheets):
    # Generating timestamp for the filename
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"AUDIT_FOR_ENGINEER_{timestamp}.xlsx"

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    workbook = writer.book

    # 1. Discrepancy Tab
    if missing_sheets or extra_sheets:
        issues = []
        for s in missing_sheets: issues.append([s, "MISSING IN TARGET"])
        for s in extra_sheets: issues.append([s, "EXTRA IN TARGET"])
        df_warn = pd.DataFrame(issues, columns=['Sheet Name', 'Discrepancy Type'])
        df_warn.to_excel(writer, sheet_name='SHEET_DISCREPANCIES', index=False)
        writer.sheets['SHEET_DISCREPANCIES'].set_tab_color('#FF9900')

    # 2. Tracking used names to prevent Excel crashes (31 char limit)
    used_names = set()

    green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

    for sheet_name, result_data in audit_results.items():
        df, action_name = result_data
        checked_rows = df[df['Audit_Status'] != 'Unknown'].copy()
        if checked_rows.empty: continue

        # --- SMART NAME TRUNCATION LOGIC ---
        # Excel limit is 31. We truncate and check for duplicates.
        base_name = sheet_name.strip()[:30]
        final_name = base_name
        counter = 1

        while final_name.lower() in used_names:
            suffix = f"_{counter}"
            # Ensure the name + suffix doesn't exceed 31
            final_name = base_name[:(31 - len(suffix))] + suffix
            counter += 1

        used_names.add(final_name.lower())
        # -----------------------------------

        checked_rows.to_excel(writer, sheet_name=final_name, index=False)
        worksheet = writer.sheets[final_name]

        # Formatting
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
    return filename


def check_sheet_consistency(target_dict, instructions_dict):
    target_sheets = set(k.strip() for k in target_dict.keys())
    instruction_sheets = set(k.strip() for k in instructions_dict.keys())
    return list(instruction_sheets - target_sheets), list(target_sheets - instruction_sheets)


def main():
    print("=" * 60)
    print("             Excel File Auditor v1.2")
    print("        Data Integrity & Consistency Engine")
    print("=" * 60)
    print("\n[STEP 1/4] Selecting files...")

    try:
        # Load Files
        target_raw, inst_raw = user_input_files()

        print("\n[STEP 2/4] Cleaning data and unmerging cells...")
        target_clean = unmerge_data(target_raw, is_instruction_file=False)
        inst_clean = unmerge_data(inst_raw, is_instruction_file=True)

        # Consistency Check
        missing, extra = check_sheet_consistency(target_clean, inst_clean)

        print("\n[STEP 3/4] Running Smart Audit...")
        final_audit_report, skipped = run_audit(target_clean, inst_clean)

        print("\n[STEP 4/4] Generating Reports...")
        # Generating a shared timestamp for all files in this session
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")

        export_to_txt(final_audit_report, f"AUDIT_RAW_{ts}.txt")
        export_summary_report(final_audit_report, missing, extra, skipped, f"AUDIT_SUMMARY_{ts}.txt")
        excel_name = export_to_excel_report(final_audit_report, missing, extra)

        # Final Notification
        messagebox.showinfo("Audit Complete",
                            f"The audit has finished successfully!\n\n"
                            f"Main Report: {excel_name}")

        print("\n" + "=" * 60)
        print("SUCCESS: Audit completed successfully.")
        print("=" * 60)

        # Friendly Exit Countdown
        print("\nClosing automatically in:", end=" ", flush=True)
        for i in range(5, 0, -1):
            print(f"{i}...", end=" ", flush=True)
            time.sleep(1)
        print("Done.")

    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: {e}")
        messagebox.showerror("Unexpected Error", f"An error occurred:\n{e}")
        # On error, we wait indefinitely so the user can read the console
        input("\nPress ENTER to close and investigate the error...")


if __name__ == "__main__":
    main()