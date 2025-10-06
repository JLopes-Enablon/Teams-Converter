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


# Function to convert 12-hour time to 24-hour format
def convert_to_24h(time_str):
    """Convert time from '8:00 AM' to '08:00' format"""
    if 'AM' in time_str or 'PM' in time_str:
        time_part = time_str.replace(' AM', '').replace(' PM', '').strip()
        hour, minute = time_part.split(':')
        hour = int(hour)
        
        if 'PM' in time_str and hour != 12:
            hour += 12
        elif 'AM' in time_str and hour == 12:
            hour = 0
            
        return f"{hour:02d}:{minute}"
    return time_str

# Regex to match event lines
# Handle both formats:
# Format 1: Mon 07/07/25 09:00 - 09:30 (24-hour)
# Format 2: Mon 10/6/25 8:00 AM - 8:30 AM (12-hour with AM/PM)
date_pattern = r'(\d{1,2}/\d{1,2}/(\d{2}|\d{4}))'
time_pattern_24h = r'(\d{1,2}:\d{2}) - (\d{1,2}:\d{2})'
time_pattern_12h = r'(\d{1,2}:\d{2} (?:AM|PM)) - (\d{1,2}:\d{2} (?:AM|PM))'

event_line_regex_24h = re.compile(r'([A-Za-z]{3}) ' + date_pattern + r' ' + time_pattern_24h)
event_line_regex_12h = re.compile(r'([A-Za-z]{3}) ' + date_pattern + r' ' + time_pattern_12h)

rows = []
for i in range(len(lines)):
    line = lines[i].strip()
    
    # Try to match 24-hour format first
    match = event_line_regex_24h.match(line)
    if match:
        raw_date = match.group(2)
        start_time = match.group(4)
        end_time = match.group(5)
        # Look for summary in previous line or next line
        summary = ""
        if i > 0 and lines[i-1].strip() and not lines[i-1].strip().startswith(('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun')):
            summary = lines[i-1].strip()
        elif i < len(lines) - 1 and lines[i+1].strip() and not lines[i+1].strip().startswith(('Location:', 'Organizer:', 'Required')):
            summary = lines[i+1].strip()
    else:
        # Try to match 12-hour format
        match = event_line_regex_12h.match(line)
        if match:
            raw_date = match.group(2)
            start_time_12h = match.group(4)
            end_time_12h = match.group(5)
            start_time = convert_to_24h(start_time_12h)
            end_time = convert_to_24h(end_time_12h)
            # Look for summary in previous line or next line
            summary = ""
            if i > 0 and lines[i-1].strip() and not lines[i-1].strip().startswith(('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun')):
                summary = lines[i-1].strip()
            elif i < len(lines) - 1 and lines[i+1].strip() and not lines[i+1].strip().startswith(('Location:', 'Organizer:', 'Required')):
                summary = lines[i+1].strip()
    
    if match and summary:
        # Normalize date format to DD/MM/YYYY
        # Handle both DD/MM/YY and DD/MM/YYYY formats
        date_parts = raw_date.split('/')
        day, month = date_parts[0], date_parts[1]
        year = date_parts[2]
        
        # Convert 2-digit year to 4-digit year
        if len(year) == 2:
            year = f"20{year}"
        
        # Format as DD/MM/YYYY for consistency
        normalized_date = f"{day.zfill(2)}/{month.zfill(2)}/{year}"
        
        print(f"Matched event: Summary='{summary}', Date='{raw_date}' -> '{normalized_date}', Start='{start_time}', End='{end_time}'")
        rows.append([summary, normalized_date, start_time, end_time])

with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Summary', 'Start Date', 'Start time', 'End Time'])
    writer.writerows(rows)

if not rows:
    print("Warning: No events were extracted. Please check the input file format and the debug output above.")
else:
    print(f'Extracted {len(rows)} events to {output_file}')
    print("\nDone! You can now open the CSV file in Excel or another spreadsheet tool.")
