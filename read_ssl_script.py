import os
import re
import json
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font

#Create folder if not exists
os.makedirs("Exports", exist_ok=True)

# Configuration
log_folder = "logs"
current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = os.path.join("Exports", f"SSLFingerprint_Report_{current_time}.xlsx")
columns = ["SrNo", "FilePath", "UserName", "LastModifiedDate"]

# Regex to extract request
request_pattern = re.compile(r"GetSSLFingerprint Request received:([^:]+):({.*})")

def process_requests(file_path, user_data, sr_counter):
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            match = request_pattern.search(line)
            if match:
                username, json_str = match.groups()

                try:
                    json_data = json.loads(json_str)
                    last_modified = json_data.get("lastModifiedDate", "")
                except json.JSONDecodeError:
                    last_modified = ""

                # Update or insert
                if username in user_data:
                    user_data[username]["LastModifiedDate"] = last_modified
                else:
                    user_data[username] = {
                        "SrNo": sr_counter,
                        "FilePath": file_path,
                        "UserName": username,
                        "LastModifiedDate": last_modified
                    }
                    sr_counter += 1
    return sr_counter

def write_to_excel(user_data, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Fingerprint Requests"

    # Header
    for col_idx, col in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col)
        cell.font = Font(bold=True)

    # Data rows
    for idx, user in enumerate(user_data.values(), start=2):
        for col_idx, col in enumerate(columns, 1):
            ws.cell(row=idx, column=col_idx, value=user.get(col, ""))

    wb.save(output_file)
    print(f"Report saved to {output_file}")

def main():
    user_data = {}
    sr_counter = 1

    for file_name in os.listdir(log_folder):
        if file_name.endswith(".txt"):
            file_path = os.path.join(log_folder, file_name)
            sr_counter = process_requests(file_path, user_data, sr_counter)

    write_to_excel(user_data, output_file)

if __name__ == "__main__":
    main()
