## Calendar Event Extractor

This script extracts event information from a Teams calendar text export and saves it as a CSV file for easy use in Excel or other tools.

### How to Use

1. Place the script (`extract_calendar_events.py`) and your exported calendar text file (e.g., `Team Cal2.txt`) in the same folder.
2. Run the script:
   ```sh
   python3 extract_calendar_events.py
   ```
3. When prompted, enter the name of your input file (or press Enter to use the default) and the desired output CSV filename.
4. The script will extract all events and save them in the CSV file you specified.

### Output Format

The CSV will have the following columns:

- Summary
- Start Date
- Start time
- End Time

Each row represents a calendar event.

### Notes
- The script automatically handles most text encoding issues.
- Only events with a line format like `Event Name` followed by `Mon 07/07/25 09:00 - 09:30` are extracted.

---
Created by GitHub Copilot
