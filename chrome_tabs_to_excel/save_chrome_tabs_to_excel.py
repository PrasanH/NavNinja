import time
import pandas as pd
import pyautogui
import pyperclip
import yaml
import os
from pathlib import Path

# Note: PyAutoGUI has a fail-safe feature - moving mouse to screen corner will abort
# Keep it enabled for safety
# If you want to disable it: pyautogui.FAILSAFE = False

# IMPORTANT: For writing to OPEN Excel files, install xlwings:
#   pip install xlwings pyyaml
# If xlwings is not installed, the file must be closed before writing.


# Load configuration from config.yaml
def load_config():
    """Load configuration from config.yaml file."""
    config_path = Path(__file__).parent / "config.yaml"

    if not config_path.exists():
        print(f"ERROR: Configuration file not found at {config_path}")
        print("Please create a config.yaml file in the same directory.")
        exit(1)

    try:
        with open(config_path, "r") as f:
            config = yaml.safe_load(f)
        return config
    except Exception as e:
        print(f"ERROR: Failed to load config.yaml: {e}")
        exit(1)


# Load configuration
CONFIG = load_config()

# Extract configuration values
EXCEL_FILE_PATH = CONFIG["excel"]["file_path"]
SHEET_NAME = CONFIG["excel"]["sheet_name"]
COLUMN_NAME = CONFIG["excel"]["column_name"]
START_ROW = CONFIG["excel"]["start_row"]

TAB_SWITCH_DELAY = CONFIG["capture"]["tab_switch_delay"]
COPY_DELAY = CONFIG["capture"]["copy_delay"]
MAX_TABS = CONFIG["capture"]["max_tabs"]
COUNTDOWN_SECONDS = CONFIG["capture"]["countdown_seconds"]


def get_all_chrome_windows():
    """
    Get all Chrome windows with their titles.
    Returns list of tuples (window_number, title).
    """
    try:
        import pygetwindow as gw

        chrome_windows = gw.getWindowsWithTitle("Chrome")
        if chrome_windows:
            windows_info = []
            for idx, window in enumerate(chrome_windows, 1):
                title = window.title.replace(" - Google Chrome", "")
                if len(title) > 60:
                    title = title[:57] + "..."
                windows_info.append((idx, title))
            return windows_info
        return []
    except ImportError:
        print("\nNote: pygetwindow not installed.")
        return [(1, "Current Chrome Window")]
    except Exception as e:
        print(f"Note: Could not detect Chrome windows: {e}")
        return [(1, "Current Chrome Window")]


def select_windows_to_capture(windows_info):
    """
    Let user select which Chrome windows to capture from.
    Returns list of window numbers to capture.
    """
    if not windows_info:
        print("No Chrome windows found!")
        return []

    print("\n" + "=" * 60)
    print("Available Chrome Windows:")
    print("=" * 60)

    for idx, title in windows_info:
        print(f"{idx}. {title}")

    print("\nOptions:")
    print("  - Enter window number(s) separated by commas (e.g., 1,3,5)")
    print("  - Enter 'all' to capture from all windows")
    print("  - Press Enter to capture from window 1 only")
    print("=" * 60)

    while True:
        try:
            choice = input("\nYour choice: ").strip()

            if choice == "":
                return [1]

            if choice.lower() == "all":
                return [idx for idx, _ in windows_info]

            # Parse comma-separated numbers
            selected = []
            for num in choice.split(","):
                num = num.strip()
                if num.isdigit():
                    window_num = int(num)
                    if 1 <= window_num <= len(windows_info):
                        selected.append(window_num)
                    else:
                        print(f"Invalid window number: {window_num}")
                        selected = []
                        break

            if selected:
                print(
                    f"\nSelected {len(selected)} window(s): {', '.join(map(str, selected))}"
                )
                return selected

        except KeyboardInterrupt:
            print("\nCancelled.")
            return []


def get_current_tab_url():
    """
    Get URL from the current Chrome tab using keyboard shortcuts.

    Returns:
        URL string or None if failed
    """
    try:
        # Select address bar (Ctrl+L)
        pyautogui.hotkey("ctrl", "l")
        time.sleep(0.2)

        # Copy URL (Ctrl+C)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(COPY_DELAY)

        # Get URL from clipboard
        url = pyperclip.paste()

        # Press Escape to deselect address bar
        pyautogui.press("escape")
        time.sleep(0.1)

        if url and url.strip():
            return url.strip()
        else:
            print(f"    Warning: Clipboard was empty or whitespace")
            return None

    except pyautogui.FailSafeException as e:
        print(f"    PyAutoGUI Fail-Safe triggered (mouse moved to corner): {e}")
        raise KeyboardInterrupt("Fail-safe triggered")
    except Exception as e:
        print(f"    Error getting URL: {type(e).__name__}: {e}")
        return None


def switch_to_next_tab():
    """Switch to the next Chrome tab."""
    pyautogui.hotkey("ctrl", "tab")
    time.sleep(TAB_SWITCH_DELAY)


def capture_tabs_from_window(window_num=1, total_windows=1):
    """
    Capture URLs from tabs in the currently focused Chrome window.

    Args:
        window_num: Current window number (for display)
        total_windows: Total number of windows being captured

    Returns:
        List of URLs from this window
    """
    if total_windows > 1:
        print(f"\n{'='*50}")
        print(f"Window {window_num} of {total_windows}")
        print(f"{'='*50}")

    print("\nCapturing tabs from this window...")

    urls = []
    seen_urls = set()

    # Get URL from first tab
    try:
        first_url = get_current_tab_url()
        if first_url:
            urls.append(first_url)
            seen_urls.add(first_url)
            print(f"  Tab 1: Captured")
        else:
            print(f"  Tab 1: Failed to get URL (returned None)")
            return urls
    except Exception as e:
        print(f"  Tab 1: Error - {e}")
        return urls

    # Capture remaining tabs
    for i in range(1, MAX_TABS):
        try:
            switch_to_next_tab()
            url = get_current_tab_url()

            if not url:
                print(f"  Tab {i+1}: Failed to get URL, stopping...")
                break

            # Check if we've cycled back to the first tab
            if url == first_url:
                print(f"  Completed cycle - returned to first tab")
                break

            # Check for duplicate
            if url in seen_urls:
                print(f"  Tab {i+1}: Duplicate detected, stopping...")
                break

            urls.append(url)
            seen_urls.add(url)
            print(f"  Tab {i+1}: Captured")

            # Safety check
            if i >= MAX_TABS - 1:
                print(f"  Reached maximum tab limit ({MAX_TABS})")
                break
        except KeyboardInterrupt:
            print(f"\n  Capture interrupted by user at tab {i+1}")
            raise
        except Exception as e:
            print(f"  Tab {i+1}: Error - {e}")
            print(f"  Stopping capture after {len(urls)} tabs")
            break

    print(f"  Captured {len(urls)} URLs from this window")
    return urls


def capture_chrome_tabs(selected_windows):
    """
    Capture URLs from selected Chrome windows.

    Args:
        selected_windows: List of window numbers to capture from

    Returns:
        List of all URLs from selected windows
    """
    num_windows = len(selected_windows)

    print("\n" + "=" * 60)
    print("Starting Chrome tab capture...")
    print("=" * 60)
    print(f"\nCapturing from {num_windows} window(s)")
    print("\nIMPORTANT:")
    print("- Do NOT switch windows or tabs manually during capture")
    print("- Do NOT press any keys during capture")
    print("- Do NOT move mouse to screen corners (this triggers fail-safe)")
    print("- The script will automatically cycle through tabs")

    if num_windows > 1:
        print(f"- You'll be prompted to focus each window")

    print()

    all_urls = []

    # Capture from each selected window
    for idx, window_num in enumerate(selected_windows, 1):
        try:
            print(f"\n--- WINDOW {window_num} ({idx} of {num_windows}) ---")
            print("INSTRUCTIONS:")
            print(f"  1. After pressing ENTER, you'll have {COUNTDOWN_SECONDS} seconds")
            print(f"  2. Use that time to click on Chrome window #{window_num}")
            print("  3. Keep Chrome window focused (don't click anywhere else)")
            input("\nPress ENTER to start the countdown...")

            print(f"\nSwitching to Chrome in:")
            for i in range(COUNTDOWN_SECONDS, 0, -1):
                print(f"  {i}...")
                time.sleep(1)
            print("  >> Starting capture NOW! <<")

            window_urls = capture_tabs_from_window(idx, num_windows)
            all_urls.extend(window_urls)
        except KeyboardInterrupt:
            print(f"\n\nCapture interrupted at window {window_num}")
            if all_urls:
                print(f"Captured {len(all_urls)} URLs so far from {idx - 1} window(s)")
                user_input = input(
                    "\nDo you want to save the URLs captured so far? (y/n): "
                )
                if user_input.lower() != "y":
                    raise
            else:
                raise

    # Remove duplicates while preserving order
    seen = set()
    unique_urls = []
    for url in all_urls:
        if url not in seen:
            seen.add(url)
            unique_urls.append(url)

    duplicates_removed = len(all_urls) - len(unique_urls)

    print(f"\n{'='*60}")
    print(f"Successfully captured {len(unique_urls)} unique URLs")
    if duplicates_removed > 0:
        print(f"({duplicates_removed} duplicate(s) removed)")
    print(f"{'='*60}\n")

    return unique_urls


def write_urls_to_excel(urls, file_path, sheet_name, column_name, start_row):
    """
    Write URLs to Excel file (works even if file is open).

    Args:
        urls: List of URLs to write
        file_path: Path to Excel file
        sheet_name: Sheet name
        column_name: Column name to write URLs
        start_row: Starting Excel row number (1-indexed)
    """
    print(f"Writing {len(urls)} URLs to DataFrame...")

    # Try xlwings first (works with open files)
    try:
        import xlwings as xw

        print("Using xlwings (can write to open files)...")
        return write_with_xlwings(urls, file_path, sheet_name, column_name, start_row)
    except ImportError:
        print("xlwings not installed. Using pandas (file must be closed)...")
        return write_with_pandas(urls, file_path, sheet_name, column_name, start_row)
    except Exception as e:
        print(f"xlwings failed: {e}")
        print("Falling back to pandas method...")
        return write_with_pandas(urls, file_path, sheet_name, column_name, start_row)


def write_with_xlwings(urls, file_path, sheet_name, column_name, start_row):
    """Write URLs using xlwings (works with open files)."""
    import xlwings as xw
    import os

    try:
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"File doesn't exist. Creating new file...")
            # Create new file with pandas first
            df = pd.DataFrame({column_name: urls})
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
            print(f"✓ Created file and wrote {len(urls)} URLs")
            return True

        # Open or connect to existing workbook
        try:
            # Try to get already open workbook
            wb = xw.Book(file_path)
            print("Connected to open Excel file")
        except:
            # Open the file
            wb = xw.Book(file_path)
            print("Opened Excel file")

        # Get or create sheet
        if sheet_name in [sheet.name for sheet in wb.sheets]:
            ws = wb.sheets[sheet_name]
        else:
            ws = wb.sheets.add(sheet_name)
            print(f"Created new sheet: {sheet_name}")

        # Find or create column
        # Get header row
        headers = ws.range("A1").expand("right").value
        if not isinstance(headers, list):
            headers = [headers] if headers else []

        # Find column index
        if column_name in headers:
            col_idx = headers.index(column_name) + 1
        else:
            # Add new column
            col_idx = len(headers) + 1
            ws.range((1, col_idx)).value = column_name
            print(f"Added column: {column_name}")

        # Check for existing data
        rows_with_data = []
        for i, url in enumerate(urls):
            row = start_row + i
            cell_value = ws.range((row, col_idx)).value
            if cell_value and str(cell_value).strip():
                rows_with_data.append(row)

        if rows_with_data:
            print(f"\n❌ ERROR: Existing data found in target row range!")
            print(
                f"\nRows {start_row}-{start_row + len(urls) - 1} already contain data in column '{column_name}'"
            )
            print(
                f"Rows with existing data: {rows_with_data[:5]}{'...' if len(rows_with_data) > 5 else ''}"
            )

            # Find next available empty row
            print(f"\nSearching for next available empty row...")
            max_row = ws.range((ws.cells.last_cell.row, col_idx)).end("up").row
            next_empty_row = max_row + 1

            # Verify it's actually empty
            for check_row in range(next_empty_row, next_empty_row + len(urls)):
                if ws.range((check_row, col_idx)).value:
                    next_empty_row = check_row + 1

            print(f"\n💡 SUGGESTION: Next available empty row is {next_empty_row}")
            print(
                f"   This would write to rows {next_empty_row}-{next_empty_row + len(urls) - 1}"
            )

            user_choice = input("\nUse this row? (y/n): ").strip().lower()
            if user_choice == "y":
                start_row = next_empty_row
                print(f"✓ Using row {start_row}")
            else:
                print("\nPlease either:")
                print(
                    f"  1. Clear the data in rows {start_row}-{start_row + len(urls) - 1}"
                )
                print(f"  2. Change START_ROW in config.yaml")
                print(f"  3. Run the script again")
                return False

        # Write URLs
        for i, url in enumerate(urls):
            row = start_row + i
            ws.range((row, col_idx)).value = url

        # Save
        wb.save()
        print(f"\n✓ Successfully wrote {len(urls)} URLs to '{file_path}'")
        print(f"  Sheet: {sheet_name}")
        print(f"  Column: {column_name}")
        print(f"  Rows: {start_row} to {start_row + len(urls) - 1}")

        return True

    except Exception as e:
        print(f"Error with xlwings: {e}")
        return False


def write_with_pandas(urls, file_path, sheet_name, column_name, start_row):
    """Write URLs using pandas (file must be closed)."""
    try:
        # Try to read existing Excel file
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"Opened existing sheet '{sheet_name}'")
        except FileNotFoundError:
            # Create new DataFrame if file doesn't exist
            df = pd.DataFrame()
            print(f"Creating new file at '{file_path}'")
        except ValueError:
            # Sheet doesn't exist, create new DataFrame
            df = pd.DataFrame()
            print(f"Creating new sheet '{sheet_name}'")

        # Ensure column exists
        if column_name not in df.columns:
            df[column_name] = None

        # Convert start_row to DataFrame index (Excel row N = pandas index N-2)
        start_idx = start_row - 2

        # Check if rows already contain data
        end_idx = start_idx + len(urls)
        rows_with_data = []

        for i in range(len(urls)):
            target_idx = start_idx + i
            # Check if this row exists and has data in the column
            if target_idx < len(df):
                cell_value = df.at[target_idx, column_name]
                # Check if cell has data (not None, not NaN, not empty string)
                if pd.notna(cell_value) and str(cell_value).strip() != "":
                    excel_row = target_idx + 2  # Convert back to Excel row number
                    rows_with_data.append(excel_row)

        # If any rows have existing data, stop with error
        if rows_with_data:
            print(f"\n❌ ERROR: Existing data found in target row range!")
            print(
                f"\nRows {start_row}-{start_row + len(urls) - 1} already contain data in column '{column_name}'"
            )
            print(
                f"Rows with existing data: {rows_with_data[:5]}{'...' if len(rows_with_data) > 5 else ''}"
            )

            # Find next available empty row
            print(f"\nSearching for next available empty row...")
            next_empty_row = start_row

            # Search for first empty range that can fit all URLs
            for search_row in range(
                2, len(df) + 1000
            ):  # Search up to 1000 rows beyond current data
                search_idx = search_row - 2
                is_range_empty = True

                for i in range(len(urls)):
                    check_idx = search_idx + i
                    if check_idx < len(df):
                        cell_value = (
                            df.at[check_idx, column_name]
                            if check_idx in df.index
                            else None
                        )
                        if pd.notna(cell_value) and str(cell_value).strip() != "":
                            is_range_empty = False
                            break

                if is_range_empty:
                    next_empty_row = search_row
                    break

            print(f"\n💡 SUGGESTION: Next available empty row is {next_empty_row}")
            print(
                f"   This would write to rows {next_empty_row}-{next_empty_row + len(urls) - 1}"
            )

            user_choice = input("\nUse this row? (y/n): ").strip().lower()
            if user_choice == "y":
                start_row = next_empty_row
                start_idx = start_row - 2
                print(f"✓ Using row {start_row}")
            else:
                print("\nPlease either:")
                print(
                    f"  1. Clear the data in rows {start_row}-{start_row + len(urls) - 1}"
                )
                print(f"  2. Change START_ROW in config.yaml")
                print(f"  3. Run the script again")
                return False

        # Ensure DataFrame is large enough
        if start_idx + len(urls) > len(df):
            # Extend DataFrame
            additional_rows = (start_idx + len(urls)) - len(df)
            empty_df = pd.DataFrame(index=range(additional_rows))
            df = pd.concat([df, empty_df], ignore_index=True)

        # Write URLs to DataFrame
        for i, url in enumerate(urls):
            df.at[start_idx + i, column_name] = url

        print(f"Writing {len(urls)} URLs to DataFrame...")

        # Save to Excel
        try:
            # Check if file exists
            import os

            file_exists = os.path.exists(file_path)

            if file_exists:
                # Read existing workbook
                from openpyxl import load_workbook

                book = load_workbook(file_path)

                # Create writer with existing workbook
                with pd.ExcelWriter(
                    file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
                ) as writer:
                    writer.book = book
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Updated existing file")
            else:
                # Create new file
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Created new file")

        except Exception as excel_error:
            print(f"Error during Excel write: {excel_error}")
            # Fallback: try simple write
            print("Trying alternative write method...")
            df.to_excel(file_path, sheet_name=sheet_name, index=False)

        print(f"\n✓ Successfully wrote {len(urls)} URLs to '{file_path}'")
        print(f"  Sheet: {sheet_name}")
        print(f"  Column: {column_name}")
        print(f"  Rows: {start_row} to {start_row + len(urls) - 1}")

        return True

    except PermissionError:
        print(f"\nError: Permission denied. Is the Excel file open?")
        print(f"Please close '{file_path}' and try again.")
        return False
    except Exception as e:
        print(f"\nError writing to Excel: {e}")
        return False


def main():
    """Main function to capture Chrome tabs and save to Excel."""
    print("=" * 60)
    print("Chrome Tabs to Excel Exporter")
    print("=" * 60)
    print(f"\nConfiguration (from config.yaml):")
    print(f"  Excel File: {EXCEL_FILE_PATH}")
    print(f"  Sheet: {SHEET_NAME}")
    print(f"  Column: {COLUMN_NAME}")
    print(f"  Starting Row: {START_ROW}")
    print()

    # Get Chrome windows
    windows_info = get_all_chrome_windows()

    if not windows_info:
        print("\nError: No Chrome windows found!")
        print("Please open Chrome and try again.")
        return

    # Let user select which windows to capture from
    selected_windows = select_windows_to_capture(windows_info)

    if not selected_windows:
        print("\nNo windows selected. Exiting.")
        return

    # Capture URLs from selected Chrome windows
    urls = capture_chrome_tabs(selected_windows)

    if not urls:
        print("\nNo URLs captured. Exiting.")
        return

    # Display captured URLs
    print("\nCaptured URLs:")
    for i, url in enumerate(urls, 1):
        # Truncate long URLs for display
        display_url = url if len(url) <= 80 else url[:77] + "..."
        print(f"  {i}. {display_url}")

    print(f"\nTotal URLs captured: {len(urls)}")

    # Write to Excel
    print(f"\nWriting to Excel...")
    print(f"Target: {EXCEL_FILE_PATH}")
    print(f"Sheet: {SHEET_NAME}, Column: {COLUMN_NAME}, Starting at row: {START_ROW}")

    success = write_urls_to_excel(
        urls, EXCEL_FILE_PATH, SHEET_NAME, COLUMN_NAME, START_ROW
    )

    if success:
        print("\n" + "=" * 60)
        print("Done! URLs saved successfully.")
        print("=" * 60)
    else:
        print("\nFailed to save URLs to Excel.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nScript interrupted by user.")
    except Exception as e:
        print(f"\n\nUnexpected error: {e}")
        import traceback

        traceback.print_exc()
