print("🚨 Script started.")
import sys
import pandas as pd
import os
import re
import datetime
import zipfile
from collections import defaultdict

base_dir = os.path.dirname(os.path.abspath(__file__))
log_path = r"C:\Users\w10itasv\Documents\PricingToolPY2\script_log.txt"
log_file = open(log_path, "w", encoding="utf-8")

def log(msg):
    print(msg)
    sys.stdout.flush()
    log_file.write(msg + "\n")
    log_file.flush()

log("🚀 SCRIPT STARTED")
log(f"📂 Script location: {base_dir}")
log(f"📂 Current working dir: {os.getcwd()}")

def normalize(text):
    return re.sub(r'\s+', ' ', str(text).strip()).upper()

def get_safe_output_path(base_dir, base_name, input_base):
    today_str = datetime.datetime.now().strftime("%Y%m%d")
    rev = 1
    short_input = input_base[:10]
    while True:
        filename = f"{base_name}-{short_input}-{today_str}-rev{rev}.csv"
        full_path = os.path.join(base_dir, filename)
        if not os.path.exists(full_path):
            return full_path
        rev += 1

def safe_open_file(operation, file_path, mode="read", retries=3):
    for attempt in range(1, retries + 1):
        try:
            return operation()
        except PermissionError:
            print(f"[Attempt {attempt}] Cannot {mode} file: {file_path}")
            if attempt < retries:
                input("Please close the file and press Enter to retry...")
            else:
                raise RuntimeError(f"Failed to {mode} '{file_path}' after {retries} attempts.")

def load_excel_file(path, mode="read", retries=3, header=None, dtype=None, sheet_name=0):
    return safe_open_file(lambda: pd.read_excel(path, header=header, dtype=dtype, sheet_name=sheet_name), path, mode, retries)

def load_text_file_lines(path, retries=3):
    return safe_open_file(lambda: open(path, "r", encoding="utf-8").readlines(), path, "read", retries)

def write_text_file(path, content_lines, retries=3):
    def operation():
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.writelines(content_lines)
    return safe_open_file(operation, path, "write", retries)

def process_file(pricelist_path, config_df, sample_dir, output_base_dir):
    input_filename = os.path.splitext(os.path.basename(pricelist_path))[0]
    output_dir = os.path.join(output_base_dir, input_filename)
    os.makedirs(output_dir, exist_ok=True)

    field_map = defaultdict(list)
    constants = defaultdict(dict)
    for _, row in config_df.iterrows():
        dest = row['Dest_table']
        field = row['Dest_field']
        src = row['Source']
        if src == "Constant":
            raw_val = row['Constant_value']
            constants[dest][field] = "" if pd.isna(raw_val) else str(raw_val).strip()
        else:
            field_map[dest].append((normalize(src), field))

    excel_file = pd.ExcelFile(pricelist_path)
    if "Load_Sheet" not in excel_file.sheet_names:
        print(f"❌ Sheet 'Load_Sheet' not found in {pricelist_path}. Skipping.")
        log(f"❌ Sheet 'Load_Sheet' not found in {pricelist_path}. Skipping.")
        sys.stdout.flush()
        return
    selected_sheet = "Load_Sheet"

    df = load_excel_file(pricelist_path, header=1, dtype=str, sheet_name=selected_sheet).fillna("")
    normalized_input_columns = {normalize(col): col for col in df.columns}

    for table in config_df['Dest_table'].unique():
        sample_file = os.path.join(sample_dir, f"1-{table}.csv")
        if not os.path.exists(sample_file):
            print(f"Sample file for '{table}' not found. Skipping.")
            log(f"Sample file for '{table}' not found. Skipping.")
            sys.stdout.flush()
            continue

        lines = load_text_file_lines(sample_file)
        header_line = lines[0].strip()
        format_line = lines[1].strip()
        headers = [h.strip().lstrip('\ufeff') for h in header_line.split(";")]
        length_defs = format_line.split(";")
        max_lengths = {}
        for i, val in enumerate(length_defs):
            match = re.match(r"\((\d+)\)", val)
            if match:
                max_lengths[headers[i]] = int(match.group(1))

        seen_keys = set() if table.lower() in ("pricelist", "baseprice") else None
        output_rows = []
        for idx, row in df.iterrows():
            pricing_type = str(row.get("PricingType", "")).strip().lower()

            combo_key = (
                str(row.get("PriceList", "")).strip(),
                str(row.get("ACCOUNT NO.", "")).strip(),
                str(row.get("CUST. ITEM", "")).strip()
            )

            if table.lower() == "baseprice":
                itno_src_field = next((src for src, dest in field_map['baseprice'] if dest == 'ITNO'), None)
                if itno_src_field and itno_src_field in normalized_input_columns:
                    itno_col = normalized_input_columns[itno_src_field]
                    itno_val = str(row.get(itno_col, "")).strip()
                    combo_key += (itno_val,)

            if seen_keys is not None and pricing_type in ("graduated", "matrix"):
                if combo_key in seen_keys:
                    continue
                seen_keys.add(combo_key)

            out_row = []
            for field in headers:
                if field in constants[table]:
                    value = constants[table][field]
                elif any(dest_field == field for _, dest_field in field_map[table]):
                    match_src = next((src for src, dest_field in field_map[table] if dest_field == field), None)
                    if match_src and match_src in normalized_input_columns:
                        source_col = normalized_input_columns[match_src]
                        raw_val = row.get(source_col, "")
                        value = "" if pd.isna(raw_val) else str(raw_val).strip()
                    else:
                        value = ""
                else:
                    value = ""

                if field in max_lengths and len(value) > max_lengths[field]:
                    print(f"Warning: Row {idx + 3}, field '{field}' exceeds max length ({max_lengths[field]}): {value}")
                    sys.stdout.flush()
                    value = value[:max_lengths[field]]

                out_row.append(value)

            output_rows.append(out_row)

        output_file = get_safe_output_path(output_dir, table, input_filename)
        output_content = [
            header_line + "\n",
            format_line + "\n"
        ] + [";".join([x if x is not None else "" for x in row]) + "\n" for row in output_rows]

        write_text_file(output_file, output_content)
        print(f"✔ Output written: {os.path.basename(output_file)} ({len(output_rows)} rows, {len(df) - len(output_rows)} skipped)")
        log(f"✔ Output written: {os.path.basename(output_file)} ({len(output_rows)} rows, {len(df) - len(output_rows)} skipped)")
        sys.stdout.flush()

    # Zip output
    zip_path = os.path.join(output_base_dir, f"{input_filename}.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(output_dir):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.relpath(full_path, start=output_dir)
                zipf.write(full_path, arcname)
    print(f"✅ Zipped: {zip_path}")
    log(f"✅ Zipped: {zip_path}")
    sys.stdout.flush()

def main():
    print("🟢 Starting script...")
    sys.stdout.flush()

    input_folder = "PriceList Input"
    config_path = "configuration.xlsx"
    sample_dir = "sample"
    output_dir = "output"

    os.makedirs(output_dir, exist_ok=True)

    config_df = load_excel_file(config_path, header=0, dtype=str)
    config_df.columns = [str(col).strip() for col in config_df.columns]
    config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

    for file_name in os.listdir(input_folder):
        if file_name.lower().endswith(('.xls', '.xlsx')):
            full_path = os.path.join(input_folder, file_name)
            process_file(full_path, config_df, sample_dir, output_dir)

    print("✅ Script completed successfully.")
    log("✅ Script completed successfully.")
    sys.stdout.flush()

print("⚠️ Entered main block.")
log("⚠️ Entered main block.")

if __name__ == "__main__":
    main()
print("📦 About to enter main block.")
log("📦 About to enter main block.")
sys.stdout.flush()
