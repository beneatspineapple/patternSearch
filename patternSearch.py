import fitz  # PyMuPDF
import os
import re
import csv
from tkinter import Tk, filedialog
from datetime import datetime
from tqdm import tqdm
import sys
import shutil

def extract_date(text):
    labels = [r"statement\s*closing\s*date", r"statement\s*end\s*date", r"closing\s*date:", r"statement\s*close\s*date"]
    numeric_date_regex = r'([0-9]{2}/[0-9]{2}/([0-9]{2,4}))'
    written_date_regex = r'([A-Za-z]+ +[0-9]{1,2}, +[0-9]{4})'

    lines = text.splitlines()
    for i, line in enumerate(lines):
        for label in labels:
            label_regex = re.compile(label, re.IGNORECASE)
            match_label = label_regex.search(line)
            if match_label:
                idx = match_label.start()
                end_idx = match_label.end()
                before = ""
                if idx >= 8:
                    before = line[idx-8:idx]
                else:
                    before = line[:idx]
                    if i > 0:
                        before = lines[i-1][-8+idx:] + before

                # Extract the date from the text after the matched label
                line_after_label = line[end_idx:]
                match = re.search(numeric_date_regex, line_after_label)
                if not match:
                    match = re.search(written_date_regex, line_after_label)
                if match:
                    date = match.group(1)
                    year = match.group(2) if len(match.groups()) > 1 else ""
                    date_format = "MM/DD/YY" if year and len(year) == 2 else ("MM/DD/YYYY" if year else "Month DD, YYYY")
                    date_pos = match.end()
                    after = line_after_label[date_pos:date_pos+30]
                    context = before + line[idx:] + after
                    return date, i+1, label, date_format, context.replace('\n', ' ').replace('\r', ' ')

                # If not found, proceed with existing logic
                combined_lines = [line[idx:]]
                for j in range(1, 4):
                    if i + j < len(lines):
                        combined_lines.append(lines[i + j])
                combined = "\n".join(combined_lines)
                match = re.search(numeric_date_regex, combined)
                if match:
                    date = match.group(1)
                    year = match.group(2)
                    date_format = "MM/DD/YY" if len(year) == 2 else "MM/DD/YYYY"
                    date_pos = match.end()
                    after = combined[date_pos:date_pos+30]
                    context = before + combined[:date_pos] + after
                    return date, i+1, label, date_format, context.replace('\n', ' ').replace('\r', ' ')
                match = re.search(written_date_regex, combined)
                if match:
                    date = match.group(1)
                    date_format = "Month DD, YYYY"
                    date_pos = match.end()
                    after = combined[date_pos:date_pos+30]
                    context = before + combined[:date_pos] + after
                    return date, i+1, label, date_format, context.replace('\n', ' ').replace('\r', ' ')
                context = before + combined
                return None, i+1, label, "", context.replace('\n', ' ').replace('\r', ' ')
    # Fallback: do NOT return a date if no label was found
    return None, "", "", "", ""

def process_pdfs(folder_path, output_csv):
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['Length', 'Filename', 'Date', 'LineNumber', 'LabelFound', 'DateFormat', 'Context'])

        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        total_files = len(pdf_files)

        for filename in tqdm(pdf_files, desc="Processing PDFs", unit="file"):
            pdf_path = os.path.join(folder_path, filename)
            try:
                doc = fitz.open(pdf_path)
                page = doc[0]
                text = page.get_text()
                doc.close()

                date, line_num, label_found, date_format, context = extract_date(text)
                file_size = os.path.getsize(pdf_path)

                writer.writerow([
                    file_size,
                    filename,
                    date if date else "NOT FOUND",
                    line_num,
                    label_found,
                    date_format if date else "NOT FOUND",
                    context
                ])
            except Exception as e:
                print(f"Error processing {filename}: {e}")


# Generate today's date string
#today_str = datetime.now().strftime("%Y%m%d")
#base_output_csv = f"search_{today_str}.csv"
#output_csv = base_output_csv
base_output_csv = f"results.csv"
output_csv = base_output_csv

# Check for existing output files and increment if needed
'''
if os.path.exists(output_csv):
    base, ext = os.path.splitext(base_output_csv)
    i = 2
    while os.path.exists(f"{base}_{i}{ext}"):
        i += 1
    output_csv = f"{base}_{i}{ext}"
'''
# Select the folder containing PDFs
try:
    folder_path = filedialog.askdirectory(
        initialdir=os.path.dirname(os.path.realpath(__file__)),
        title='Select a Directory'
    )
    if not folder_path:
        print("No folder selected. Exiting.")
        sys.exit(1)
    print(f"Selected folder: {folder_path}")

except Exception as e:
    print(f"An error occurred: {e}")

# Process the PDFs and generate the CSV
process_pdfs(folder_path, output_csv)

# Backup logic
output_filename = output_csv
backup_folder = "backup"

# Ensure backup folder exists
os.makedirs(backup_folder, exist_ok=True)

# Prepare date and base backup filename
today = datetime.now().strftime("%Y%m%d")
base_backup_filename = f"results_{today}"

# Find the next available index
index = 1
while True:
    backup_filename = f"{base_backup_filename}_{index}.csv"
    backup_path = os.path.join(backup_folder, backup_filename)
    if not os.path.exists(backup_path):
        break
    index += 1

# Copy to backup folder with unique name
shutil.copy(output_filename, backup_path)

print(f"Processing complete. Results saved to {output_csv}.")
print(f"Backup saved to {os.path.join(backup_folder, backup_filename)}")


