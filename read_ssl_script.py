import os
import re
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime

# Define folder and output file
log_folder = "logs"
current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f"Fingerprint_Report_{current_time}.xlsx"

# Column headers
columns = ["SrNo", "FilePath", "UserName", "Type", "Data"]

# Regex patterns
request_pattern = re.compile(r"GetSSLFingerprint Request received:([^:]+):({.*})")
response_pattern = re.compile(r"GetSSLFingerprint Response sent:([^:]+):({.*})")

# Processing function
def process_fingerprint_logs(file_path, sr_start=1):
    data = []
    sr_no = sr_start

    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            for log_type, pattern in [("request", request_pattern), ("response", response_pattern)]:
                match = pattern.search(line)
                if match:
                    username, json_data = match.groups()
                    data.append({
                        "SrNo": sr_no,
                        "FilePath": file_path,
                        "UserName": username,
                        "Type": log_type,
                        "Data": json_data
                    })
                    sr_no += 1
    return data

# Excel write function
def write_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Fingerprint Logs"

    # Write header
    for col_idx, col in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col)
        cell.font = Font(bold=True)

    # Write rows
    for idx, row in enumerate(data, start=2):
        for col_idx, col in enumerate(columns, 1):
            ws.cell(row=idx, column=col_idx, value=row.get(col, ""))

    wb.save(output_file)
    print(f"Report saved to {output_file}")

# Main runner
def main():
    all_data = []
    sr_counter = 1
    for file_name in os.listdir(log_folder):
        if file_name.endswith(".txt"):
            file_path = os.path.join(log_folder, file_name)
            data = process_fingerprint_logs(file_path, sr_counter)
            all_data.extend(data)
            sr_counter += len(data)

    write_to_excel(all_data, output_file)

if __name__ == "__main__":
    main()
