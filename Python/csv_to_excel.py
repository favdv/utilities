# --------------------------------------------------------------------------
# Python function: csv_to_excel.py
# Requires the following libraries: pandas and openpyxl
# 
# Purpose: 
# - Converts CSV files in a folder to an Excel spreadsheet.
# - Each CSV is converted to a separate tab.
# - The tab name is based on the CSV filename, but sanitized to avoid invalid names.
# - Data is converted to an Excel table with the same name as the tab name.
# 
# Usage:
# - By default, the script converts all CSV files in the same location as the Python file
#   and saves the Excel file as "combined_csvs.xlsx" in the same folder.
# - The script can be executed via the command line with optional parameters:
#
# Command-line Arguments:
#  --path <directory>       : (Optional) Path to the folder containing CSV files. Defaults to the current folder.
#  --output <filename.xlsx> : (Optional) Name of the output Excel file. Defaults to "combined_csvs.xlsx".
#  --delimiter <char>       : (Optional) Delimiter used in CSV files (e.g. "|"). Defaults to a comma (",").
#  --ignoreCsvs <files>     : (Optional) Comma-separated list of CSV files to ignore (e.g., "file1.csv,file2.csv").
# 
# Example Usage:
# 1. Convert all CSVs in the current directory:
#    python csv_to_excel.py
#
# 2. Specify a custom output filename:
#    python csv_to_excel.py --output my_output.xlsx
#
# 3. Convert CSVs from a specific folder:
#    python csv_to_excel.py --path /path/to/csvs/
#
# 4. Use a different delimiter (e.g., tab-separated files):
#    python csv_to_excel.py --delimiter "\t"
#
# 5. Ignore specific CSV files:
#    python csv_to_excel.py --ignoreCsvs "ignore_this.csv,skip_this_too.csv"
#
# Restrictions:
# - The output file must be stored in the same folder as the CSV files (specified by --path), or in a subfolder or sibling folder.
# - If an invalid location is specified (e.g., an unrelated absolute path), the script will raise an error.
# - If --path and --output contain spaces, they need to be specified in quotes. 
# - The functionality currently does not support multiple or mixed delimiters across csv files.
#
# Assumptions:
# - The options --path and --output can be absolute or relative paths
# - The output filename must end with ".xlsx"; otherwise, an error will be raised.
# - Files listed in --ignoreCsvs must have the .csv extension and must be located within the directory specified by --path.  
# - It is assumed that CSV files are properly structured with headers and consistent columns.
# - If a CSV is malformed (e.g., missing headers or inconsistent columns), unexpected behavior may occur, such as errors or missing data.
# --------------------------------------------------------------------------



# Defaults (defined right at the top for easy access and modification)
DEFAULT_OUTPUT = "combined_csvs.xlsx"  # Default name for the output Excel file
DEFAULT_DELIMITER = ","  # Default delimiter for CSV files (comma-separated)
DEFAULT_IGNORE_CSVS = []  # Default list of CSV files to ignore (empty by default)

# Standard library imports
import os
import re
import sys
import argparse

# Third-party library imports
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


def verify_libraries(required_libraries: list) -> None:
    """
    Verifies that all required Python libraries are installed.
    If any library is missing, it provides an actionable message
    and exits the script gracefully.

    Args:
        required_libraries (list): List of library names to verify.
    """
    missing_libraries = []
    for library in required_libraries:
        try:
            __import__(library)  # Dynamically try importing each library
        except ImportError:
            missing_libraries.append(library)
    if missing_libraries:
        print(f"The following libraries are missing:")
        for lib in missing_libraries:
            print(f" - {lib}")
        print(f"Install them with: pip install " + " ".join(missing_libraries))
        sys.exit(1)  # Exit the script since dependencies are critical
    print(f"Libraries in place")


def sanitize_name(name: str, max_length: int = 31) -> str:
    """
    Ensures Excel tab and table names are valid by:
    - Replacing invalid characters with underscores.
    - Truncating names exceeding Excel's 31-character limit.

    Args:
        name (str): The name to sanitize.
        max_length (int): Maximum allowed length for the name.

    Returns:
        str: Sanitized and truncated name.
    """
    sanitized = re.sub(r"[^\w]", "_", name)  # Replace invalid characters
    result = (sanitized[:max_length - 3] + "...") if len(sanitized) > max_length else sanitized
    return result


def validate_and_create_output_folder(output_path: str, input_path: str) -> None:
    """
    Validates the output folder based on the following criteria:
    - If the folder doesn't exist, it is created only if it is a child or sibling of the input path.
    - If the folder exists, verifies write permissions to ensure the output file can be saved.

    Args:
        output_path (str): Full path to the output file.
        input_path (str): Full path to the input folder containing CSVs.

    Raises:
        ValueError: If the folder cannot be created.
        PermissionError: If the folder exists but is not writable.
    """
    output_folder = os.path.dirname(output_path)  # Extract folder from output path
    if not output_folder:  # If no folder is specified (e.g., "output.xlsx"), no validation is needed
        return

    # Resolve absolute paths for comparison
    output_folder_abs = os.path.abspath(output_folder)
    input_path_abs = os.path.abspath(input_path)

    # Check if the folder exists
    if not os.path.exists(output_folder_abs):
        # Allow folder creation if it's a child or sibling of the input path
        if output_folder_abs.startswith(input_path_abs) or os.path.dirname(output_folder_abs) == os.path.dirname(input_path_abs):
            print(f"Creating output folder: {output_folder_abs}")
            os.makedirs(output_folder_abs)
        else:
            raise ValueError(
                f"Error: The specified output folder does not exist and cannot be created.\n\n"
                f"Please ensure that the folder exists or:\n"
                f" - Use a child folder within the input path.\n"
                f" - Use a sibling folder to the input path.\n"
                f"\nFor example:\n"
                f" - Child folder: {os.path.join(input_path_abs, 'subfolder/output.xlsx')}\n"
                f" - Sibling folder: {os.path.join(os.path.dirname(input_path_abs), 'output.xlsx')}\n\n"
                f"Invalid folder path: {output_folder_abs}"
            )

    # Verify write permissions by attempting to create a temporary file
    try:
        temp_file = os.path.join(output_folder_abs, ".write_test.tmp")
        with open(temp_file, "w") as f:
            f.write("Write test")  # Write to test permissions
        os.remove(temp_file)  # Clean up the temporary file
    except IOError:
        raise PermissionError(
            f"Error: Cannot write to the specified output folder.\n\n"
            f"Please check the folder's permissions and ensure that you have write access.\n"
            f"Invalid folder path: {output_folder_abs}"
        )


def combine_csvs(
    path: str = None,
    output: str = DEFAULT_OUTPUT,
    delimiter: str = DEFAULT_DELIMITER,
    ignore_csvs: list = DEFAULT_IGNORE_CSVS
) -> None:
    """
    Combines multiple CSV files from a specified directory into a single Excel workbook.
    Each CSV file is placed into a separate tab with a structured table.

    Args:
        path (str): Path to the folder containing CSV files. Defaults to current directory.
        output (str): Path to the output Excel file. Defaults to DEFAULT_OUTPUT.
        delimiter (str): Delimiter used in the CSV files. Defaults to DEFAULT_DELIMITER.
        ignore_csvs (list): List of CSV files to ignore. Defaults to DEFAULT_IGNORE_CSVS.
    """
    
    path = os.path.abspath(path or os.getcwd())  # Resolve the input path
    print(f"Input directory resolved to: {path}")
    
    ignore_csvs = ignore_csvs or []  # List of files to ignore
    print(f"Files to ignore: {ignore_csvs}")

    # Validate and prepare the output folder
    output_path = os.path.abspath(output)
    print(f"Output file path resolved to: {output_path}")
    try:
        validate_and_create_output_folder(output_path, path)
    except Exception as e:
        print(f"Failed to validate or create the output folder. Error: {e}")
        return
    
    print(f"Start processing CSV files...")

    # Initialize a new Excel workbook
    wb = Workbook()
    wb.remove(wb.active)  # Remove the default sheet

    # Iterate over files in the input directory
    for filename in os.listdir(path):
        if filename.endswith(".csv") and filename not in ignore_csvs:
            try:
                file_path = os.path.join(path, filename)
                print(f"Processing file: {filename}")
                
                # Read CSV into a DataFrame
                df = pd.read_csv(file_path, delimiter=delimiter, header=0)
                if df.empty:
                    print(f"Skipping empty CSV file: {filename}")
                    continue

                # Sanitize tab and table names
                tab_name = sanitize_name(os.path.splitext(filename)[0])
                table_name = sanitize_name(tab_name, max_length=31)
                
                # Create a new sheet in the workbook
                ws = wb.create_sheet(title=tab_name)
                ws.append(list(df.columns))  # Write column headers to the sheet
                for row in df.itertuples(index=False):
                    ws.append(row)

                # Apply red color to custom headers (if any exist)
                for col_index, header in enumerate(df.columns, start=1):
                    if header.startswith("_custom_col"):
                        ws.cell(row=1, column=col_index).fill = PatternFill(
                            start_color="FF0000", end_color="FF0000", fill_type="solid"
                        )

                # Define and add the table
                table_range = f"A1:{chr(65 + len(df.columns) - 1)}{len(df) + 1}"
                table = Table(displayName=table_name, ref=table_range)
                style = TableStyleInfo(
                    name="TableStyleMedium9", showRowStripes=True, showColumnStripes=True
                )
                table.tableStyleInfo = style
                ws.add_table(table)
                print(f"Added table to sheet: {tab_name}")

            except Exception as e:
                print(f"Error processing file {filename}: {e}")

    # Save the workbook
    try:
        wb.save(output_path)
        print(f"Excel file created successfully! File saved as: {output_path}")
        print(f"Make sure you check the output. If not, it might be due to the associated CSV being misconfigured.")
    except Exception as e:
        print(f"Failed to save Excel file. Error: {e}")


if __name__ == "__main__":
    print(f"Starting process...")
    
    verify_libraries(["pandas", "openpyxl"])
    parser = argparse.ArgumentParser(description="Combine CSV files into a single Excel file.")
    parser.add_argument(
        "--path",
        type=str,
        help="Directory containing the CSV files. If not specified, defaults to the current folder."
    )
    parser.add_argument(
        "--output",
        type=str,
        default=DEFAULT_OUTPUT,
        help=f"Output Excel filename. Defaults to: {DEFAULT_OUTPUT}"
    )
    parser.add_argument(
        "--delimiter",
        type=str,
        default=DEFAULT_DELIMITER,
        help=f"CSV delimiter. Defaults to: '{DEFAULT_DELIMITER}'"
    )
    parser.add_argument(
        "--ignoreCsvs",
        type=str,
        help="Comma-separated list of CSV files to ignore (e.g., 'file1.csv,file2.csv')."
    )

    # Execute the main function
    args = parser.parse_args()
    ignore_csvs = args.ignoreCsvs.split(",") if args.ignoreCsvs else []

    combine_csvs(
        path=args.path,
        output=args.output,
        delimiter=args.delimiter,
        ignore_csvs=ignore_csvs
    )
