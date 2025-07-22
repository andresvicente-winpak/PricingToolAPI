import sys
import pandas as pd
import os
import re
import datetime
import zipfile
from collections import defaultdict

output_dir = "PricingToolPY2/Pricing Outputs"
os.makedirs(output_dir, exist_ok=True)

# Ensure log file exists
log_dir = "PricingToolPY2"
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "FixedLoadSheet_log.txt")
if not os.path.exists(log_path):
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("Log file initialized.\n")

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
    output_subdir = os.path.join(output_base_dir, input_filename)
    os.makedirs(output_subdir, exist_ok=True)

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
        return
    selected_sheet = "Load_Sheet"

    df = load_excel_file(pricelist_path, header=1, dtype=str, sheet_name=selected_sheet).fillna("")
    normalized_input_columns = {normalize(col): col for col in df.columns}

    for table in config_df['Dest_table'].unique():
        sample_file = os.path.join(sample_dir, f"1-{table}.csv")
        if not os.path.exists(sample_file):
            print(f"Sample file for '{table}' not found. Skipping.")
            continue

        lines = load_text_file_lines(sample_file)
        header_line = lines[0].strip()
        format_line = lines[1].strip()
        headers = [h.strip().lstrip('\ufeff') for h in header_line.split(";")]
        length_defs = format_line.split(";")
        max_lengths = {}
        for i, val in enumerate(length_defs):
            match = re.match(r"\\((\\d+)\\)", val)
            if match:
                max_lengths[headers[i]] = int(match.group(1))

        seen_keys = set() if table.lower() in ("pricelist", "baseprice") else None
        output_rows = []
        for idx, row in df.iterrows():
            pricing_type = str(row.get("PricingType", "")).strip().lower()
            base_price_zero = str(row.get("Base_Price_zero", "F")).strip().upper() == "T"

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

                if field == "SAPR" and table.lower() == "baseprice" and base_price_zero:
                    value = "0.0000"
                elif field == "SAPR":
                    try:
                        value = f"{float(value):.4f}"
                    except:
                        value = value

                if field in max_lengths and len(value) > max_lengths[field]:
                    print(f"Warning: Row {idx + 3}, field '{field}' exceeds max length ({max_lengths[field]}): {value}")
                    value = value[:max_lengths[field]]

                out_row.append(value)

            output_rows.append(out_row)

        output_file = get_safe_output_path(output_subdir, table, input_filename)
        output_content = [
            header_line + "\n",
            format_line + "\n"
        ] + [";".join([x if x is not None else "" for x in row]) + "\n" for row in output_rows]

        write_text_file(output_file, output_content)
        print(f"✔ Output written: {os.path.basename(output_file)} ({len(output_rows)} rows, {len(df) - len(output_rows)} skipped)")

    zip_path = os.path.join(output_base_dir, f"{input_filename}.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(output_subdir):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.relpath(full_path, start=output_subdir)
                zipf.write(full_path, arcname)
    print(f"✅ Zipped: {zip_path}")

def main():
    with open(log_path, "w", encoding="utf-8") as log_file:
        def log(msg):
            print(msg)
            sys.stdout.flush()
            log_file.write(msg + "\n")
            log_file.flush()

        log("🚀 Script started.")
        log(f"📂 Log path: {log_path}")

        input_folder = "PriceList Input"
        config_path = "configuration.xlsx"
        sample_dir = "sample"
        os.makedirs(output_dir, exist_ok=True)

        log(f"📁 Looking in folder: {input_folder}")
        log(f"📄 Files found: {os.listdir(input_folder)}")

        config_df = load_excel_file(config_path, header=0, dtype=str)
        config_df.columns = [str(col).strip() for col in config_df.columns]
        config_df = config_df.dropna(subset=['Dest_table', 'Dest_field'])

        for file_name in os.listdir(input_folder):
            if file_name.lower().endswith(('.xls', '.xlsx')):
                full_path = os.path.join(input_folder, file_name)
                process_file(full_path, config_df, sample_dir, output_dir)
                log(f"✅ Processed {file_name}")

        log("✅ Script completed successfully.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"❌ Unhandled exception: {str(e)}\n")
            import traceback
            traceback.print_exc(file=log_file)
        raise
