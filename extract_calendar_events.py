import re
import csv
import os

# Try to import pdfplumber if available
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

def prompt_filename(prompt_text, default):
    user_input = input(f"{prompt_text} [{default}]: ").strip()
    return user_input if user_input else default

input_file = prompt_filename("Enter the input calendar text or PDF filename", "Team Cal2.txt")
output_file = prompt_filename("Enter the output CSV filename", "Team Cal2_events.csv")

# Read lines from TXT or PDF
lines = []
file_ext = os.path.splitext(input_file)[1].lower()
if file_ext == '.pdf':
    if not pdfplumber:
        print("Error: pdfplumber is required for PDF extraction. Please install it with 'pip install pdfplumber'.")
        exit(1)
    with pdfplumber.open(input_file) as pdf:
        text = "\n".join(page.extract_text() or '' for page in pdf.pages)
        print("\n--- Debug: First 500 characters of extracted text from PDF ---\n")
        print(text[:500])
        print("\n--- Debug: First 10 lines extracted from PDF ---\n")
        lines = text.splitlines()
        for idx, l in enumerate(lines[:10]):
            print(f"{idx+1}: {l}")
        if not lines:
            print("Warning: No lines extracted from PDF. The PDF may be empty or not text-based.")
else:
    try:
        with open(input_file, encoding='utf-8') as f:
            lines = f.readlines()
    except UnicodeDecodeError:
        with open(input_file, encoding='latin-1') as f:
            lines = f.readlines()


# Regex to match event lines
# Example: Azure Daily\nMon 07/07/25 09:00 - 09:30 or Mon 21/07/2025 09:30 - 09:55
date_pattern = r'(\d{2}/\d{2}/(\d{2}|\d{4}))'
event_line_regex = re.compile(r'([A-Za-z]{3}) ' + date_pattern + r' (\d{2}:\d{2}) - (\d{2}:\d{2})')

rows = []
for i in range(len(lines) - 1):
    summary = lines[i].strip()
    match = event_line_regex.match(lines[i+1].strip())
    if match and summary:
        start_date = match.group(2)
        start_time = match.group(4)
        end_time = match.group(5)
        print(f"Matched event: Summary='{summary}', Date='{start_date}', Start='{start_time}', End='{end_time}'")
        rows.append([summary, start_date, start_time, end_time])

with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Summary', 'Start Date', 'Start time', 'End Time'])
    writer.writerows(rows)

if not rows:
    print("Warning: No events were extracted. Please check the input file format and the debug output above.")
else:
    print(f'Extracted {len(rows)} events to {output_file}')
    print("\nDone! You can now open the CSV file in Excel or another spreadsheet tool.")
