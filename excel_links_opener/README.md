# Excel Links Opener

Open URLs from an Excel file in Chrome browser, with support for batch processing and row filtering.

## Features

- ✅ Read URLs from any Excel file (.xlsx)
- ✅ Select specific sheet and column
- ✅ Filter by row range (e.g., rows 824-835)
- ✅ Batch processing with configurable batch size
- ✅ Configurable delays between batches
- ✅ YAML configuration for easy setup
- ✅ Automatic filtering of empty/invalid URLs

## Installation

1. Install required Python packages:
```bash
pip install -r requirements.txt
```

## Quick Start

1. **Configure the script:**
   - Edit `config.yaml` with your Excel file path, sheet name, and column name
   - Set the row range (or leave as `null` to read all rows)
   - Adjust batch size and delay if needed

2. **Run the script:**
```bash
python open_links_from_excel.py
```

3. **URLs will open in Chrome:**
   - First batch opens in a new window
   - Subsequent batches open in new tabs
   - Configurable delay between batches

## Configuration

Edit `config.yaml` to customize the behavior:

```yaml
excel:
  file_path: "H:\\my_excel.xlsx"
  sheet_name: "my_sheet"
  column_name: "Link"
  start_row: 824  # First row to read (null = all rows from start)
  end_row: 835    # Last row to read (null = all rows to end)

browser:
  chrome_path: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"

opening:
  batch_size: 4                 # URLs per batch
  delay_between_batches: 4      # Seconds between batches
```

### Configuration Options

**excel section:**
- `file_path`: Full path to your Excel file
- `sheet_name`: Name of the sheet containing URLs
- `column_name`: Name of the column with URLs
- `start_row`: First row to read (Excel row number, 1-indexed). Use `null` to start from beginning
- `end_row`: Last row to read (Excel row number, 1-indexed). Use `null` to read to end

**browser section:**
- `chrome_path`: Full path to Chrome executable

**opening section:**
- `batch_size`: Number of URLs to open at once (recommended: 4-6)
- `delay_between_batches`: Seconds to wait between batches (recommended: 3-5)

## How It Works

1. **Load Configuration:** Reads settings from `config.yaml`
2. **Read Excel File:** Opens the specified Excel file and sheet
3. **Extract URLs:** Gets URLs from the specified column, optionally filtered by row range
4. **Filter Empty Values:** Automatically removes empty cells and null values
5. **Open in Batches:** Opens URLs in Chrome in configurable batch sizes with delays

### Row Numbering

- Excel rows are 1-indexed (row 1 is the header)
- Data typically starts at row 2
- Example: To open rows 824-835 from Excel, set `start_row: 824` and `end_row: 835`

## Tips

- **Batch Size:** Keep batch size reasonable (4-8) to avoid overwhelming Chrome
- **Delay:** Use at least 2-3 seconds delay between batches for smooth loading
- **Row Range:** Use row range to process specific sections of large spreadsheets
- **Testing:** Start with a small row range (2-3 rows) when testing with a new file

## Troubleshooting

**"File not found" error:**
- Check that the `file_path` in config.yaml is correct
- Use double backslashes (`\\`) in Windows paths
- Or use forward slashes: `H:/CVT/Job_Hunt/Jobs.xlsx`

**"Sheet not found" error:**
- Verify the `sheet_name` matches exactly (case-sensitive)
- Check for extra spaces in the sheet name

**"Column not found" error:**
- Verify the `column_name` matches exactly (case-sensitive)
- The script will list available columns to help you

**Chrome doesn't open:**
- Check that `chrome_path` points to your Chrome installation
- Default path: `C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe`

**Invalid row range:**
- Ensure `start_row` is less than `end_row`
- Remember row 1 is the header, data starts at row 2
- Use `null` to read all rows

**No URLs found:**
- Check that the specified column contains URL data
- Verify that the row range includes data rows (not just the header)

## Requirements

- Python 3.7+
- pandas
- openpyxl
- pyyaml
- Google Chrome browser

## Examples

**Example 1: Open all URLs from a column**
```yaml
start_row: null
end_row: null
```

**Example 2: Open specific rows (Excel rows 824-835)**
```yaml
start_row: 824
end_row: 835
```

**Example 3: Open from row 100 to end**
```yaml
start_row: 100
end_row: null
```

**Example 4: Open first 50 data rows**
```yaml
start_row: 2
end_row: 51
```
