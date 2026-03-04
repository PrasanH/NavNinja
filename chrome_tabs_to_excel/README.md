# Chrome Tabs to Excel

Save URLs from all your open Chrome tabs directly to an Excel file!

## Features

- 📑 Select specific Chrome windows or capture from all windows
- 🔄 Works even when Excel file is open (with xlwings)
- ✅ Smart row detection - suggests next available empty row
- 🛡️ Prevents overwriting existing data
- ⚙️ Easy configuration via YAML file

## Installation

Install required packages:

```bash
pip install pandas openpyxl pyautogui pyperclip pyyaml xlwings pygetwindow
```

## Quick Start

1. **Edit configuration** - Open `config.yaml` and update:
   - Excel file path
   - Sheet name
   - Column name
   - Starting row number

2. **Open Chrome tabs** - Open all the tabs you want to save

3. **Run the script**:
   ```bash
   python save_chrome_tabs_to_excel.py
   ```

4. **Select windows** - Choose which Chrome windows to capture from

5. **Focus Chrome** - After pressing ENTER, you'll have 3 seconds to click on Chrome window

6. **Done!** - URLs are automatically captured and saved to Excel

## Configuration (config.yaml)

```yaml
excel:
  file_path: "H:\\my_excel.xlsx"  # Your Excel file path
  sheet_name: "my_sheet"                    # Sheet to write to  
  column_name: "Link"                        # Column for URLs
  start_row: 840                             # Starting row number

capture:
  tab_switch_delay: 0.5    # Delay between switching tabs (seconds)
  copy_delay: 0.3          # Delay after copying URL (seconds)
  max_tabs: 20             # Max tabs per window (safety limit)
  countdown_seconds: 3     # Time to focus Chrome after pressing ENTER
```

## How It Works

1. **Window Selection** - Lists all Chrome windows, you choose which to capture
2. **Tab Capture** - Cycles through tabs using Ctrl+Tab, copies URLs from address bar
3. **Duplicate Removal** - Automatically removes duplicate URLs
4. **Smart Writing** - Checks for existing data, suggests next empty row if needed
5. **Excel Update** - Writes URLs to specified location

## Tips

- ✅ Close Excel file for faster writing (or install xlwings for open file support)
- ✅ Don't touch keyboard/mouse during capture
- ✅ Keep mouse away from screen corners (triggers fail-safe)
- ✅ Update `start_row` in config.yaml for next batch

## Error Handling

- **Data exists?** - Script suggests next available empty row
- **File open?** - xlwings handles open files (pandas requires closed file)
- **Capture interrupted?** - Option to save captured URLs so far

## Troubleshooting

**Issue:** "No Chrome windows found"
- **Solution:** Make sure Chrome is running

**Issue:** "Permission denied" when writing
- **Solution:** Close Excel file or install xlwings

**Issue:** Tabs not captured correctly
- **Solution:** Increase delays in config.yaml

## Example Session

```
Available Chrome Windows:
1. LinkedIn Jobs - Data Scientist
2. Indeed - Software Engineer
3. GitHub Repositories

Your choice: 1,2

Captured 15 URLs from window 1
Captured 12 URLs from window 2
Total: 25 unique URLs

✓ Successfully wrote 25 URLs to Jobs.xlsx
  Rows: 840-864
```

## Files

- `save_chrome_tabs_to_excel.py` - Main script
- `config.yaml` - Configuration file (edit this!)
- `README.md` - This file

---

**Note:** Always backup your Excel file before running the script!
