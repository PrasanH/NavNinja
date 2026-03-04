import subprocess
import time
import pandas as pd
import yaml
from pathlib import Path


def load_config():
    """Load configuration from YAML file."""
    config_path = Path(__file__).parent / "config.yaml"

    try:
        with open(config_path, "r") as f:
            config = yaml.safe_load(f)
        return config
    except FileNotFoundError:
        print(f"Error: Configuration file not found at {config_path}")
        print(
            "Please create a config.yaml file using config.yaml.template as reference."
        )
        exit(1)
    except yaml.YAMLError as e:
        print(f"Error parsing config.yaml: {e}")
        exit(1)


def read_urls_from_excel(
    file_path, sheet_name, column_name, start_row=None, end_row=None
):
    """
    Read URLs from an Excel file.

    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet to read from
        column_name: Name of the column containing URLs
        start_row: Starting Excel row number (1-indexed, inclusive). None = start from beginning
        end_row: Ending Excel row number (1-indexed, inclusive). None = read to end

    Returns:
        List of URLs (excluding empty/null values)
    """
    try:
        # Read Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Check if column exists
        if column_name not in df.columns:
            print(f"Error: Column '{column_name}' not found in sheet '{sheet_name}'")
            print(f"Available columns: {', '.join(df.columns)}")
            return []

        # Filter by row range if specified
        # Note: Excel rows are 1-indexed, pandas DataFrame is 0-indexed
        # Row 1 in Excel is the header (row 0 in pandas after reading with header)
        # Row 2 in Excel is index 0 in pandas DataFrame
        if start_row is not None or end_row is not None:
            # Convert Excel row numbers to pandas indices
            # Excel row N corresponds to pandas index N-2 (since row 1 is header)
            start_idx = (start_row - 2) if start_row is not None else 0
            end_idx = (end_row - 1) if end_row is not None else len(df)

            # Ensure valid range
            start_idx = max(0, start_idx)
            end_idx = min(len(df), end_idx)

            if start_idx >= end_idx:
                print(
                    f"Error: Invalid row range. Start row {start_row} must be less than end row {end_row}"
                )
                return []

            print(f"Filtering rows {start_row} to {end_row} (Excel row numbers)")
            df = df.iloc[start_idx:end_idx]

        # Extract URLs from column and remove null/empty values
        urls = df[column_name].dropna().tolist()

        # Convert to strings and filter out empty strings
        urls = [
            str(url).strip()
            for url in urls
            if str(url).strip() and str(url).lower() != "nan"
        ]

        print(f"Found {len(urls)} non-empty URLs in column '{column_name}'")
        return urls

    except FileNotFoundError:
        print(f"Error: File not found at '{file_path}'")
        return []
    except ValueError as e:
        print(f"Error: Sheet '{sheet_name}' not found in the Excel file")
        print(f"Error details: {e}")
        return []
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []


def open_urls_in_batches(urls, chrome_path, batch_size, delay):
    """
    Open URLs in Chrome browser in batches.

    Args:
        urls: List of URLs to open
        chrome_path: Path to Chrome executable
        batch_size: Number of URLs to open per batch
        delay: Delay in seconds between batches
    """
    if not urls:
        print("No URLs to open!")
        return

    print(f"\nOpening {len(urls)} URLs in batches of {batch_size}...")

    # Open URLs in batches
    for i in range(0, len(urls), batch_size):
        batch = urls[i : i + batch_size]

        # For first batch, open new window; for others, open in existing window
        if i == 0:
            command = [chrome_path, "--new-window"] + batch
        else:
            command = [chrome_path] + batch

        try:
            subprocess.Popen(command)
            print(f"Opened batch {i//batch_size + 1}: {len(batch)} URLs")

            # Add delay between batches (except for last batch)
            if i + batch_size < len(urls):
                time.sleep(delay)
        except Exception as e:
            print(f"Error opening batch {i//batch_size + 1}: {e}")


def main():
    """Main function to read URLs from Excel and open them in Chrome."""
    # Load configuration
    config = load_config()

    # Extract configuration values
    excel_config = config["excel"]
    browser_config = config["browser"]
    opening_config = config["opening"]

    file_path = excel_config["file_path"]
    sheet_name = excel_config["sheet_name"]
    column_name = excel_config["column_name"]
    start_row = excel_config.get("start_row")
    end_row = excel_config.get("end_row")
    chrome_path = browser_config["chrome_path"]
    batch_size = opening_config["batch_size"]
    delay = opening_config["delay_between_batches"]
    max_limit = opening_config.get("max_limit")

    print("=" * 50)
    print("Excel URL Opener for Chrome")
    print("=" * 50)
    print(f"\nReading from:")
    print(f"  File: {file_path}")
    print(f"  Sheet: {sheet_name}")
    print(f"  Column: {column_name}")
    if start_row is not None or end_row is not None:
        row_info = f"  Rows: {start_row or 'start'} to {end_row or 'end'}"
        print(row_info)
    else:
        print(f"  Rows: All")
    print()

    # Read URLs from Excel
    urls = read_urls_from_excel(file_path, sheet_name, column_name, start_row, end_row)

    # Apply max_limit if specified
    if urls and max_limit is not None and max_limit > 0:
        if len(urls) > max_limit:
            print(
                f"\nApplying max_limit: Opening only {max_limit} out of {len(urls)} URLs"
            )
            urls = urls[:max_limit]

    # Open URLs in batches
    if urls:
        open_urls_in_batches(urls, chrome_path, batch_size, delay)
        print("\nDone!")
    else:
        print("\nNo URLs found or error occurred.")


if __name__ == "__main__":
    main()
