import os
import shutil
import pandas as pd
from pathlib import Path
import logging
from datetime import datetime

# Set up logging
log_file = 'rename_folder.log'
logging.basicConfig(filename=log_file, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    filemode='a')  # 'a' means append mode


def rename_folders():
    # Define paths
    source_dir = Path('./1-BEFORE')
    target_dir = Path('./2-AFTER')
    excel_file = 'RenameFolder.xlsx'

    logging.info("Starting folder renaming operation")

    # Check if target directory is empty
    if any(target_dir.iterdir()):
        logging.error("The ./2-AFTER folder is not empty")
        print(
            "Error: The ./2-AFTER folder is not empty. Please empty it before proceeding.")
        return

    # Read Excel file
    try:
        df = pd.read_excel(excel_file, sheet_name='rename')
    except FileNotFoundError:
        logging.error(f"The file {excel_file} was not found")
        print(f"Error: The file {excel_file} was not found.")
        return
    except Exception as e:
        logging.error(f"Error reading the Excel file: {e}")
        print(f"Error reading the Excel file: {e}")
        return

    # Get list of folders in source directory
    source_folders = set(f.name for f in source_dir.iterdir() if f.is_dir())
    excel_folders = set(df['Folder Lama'])

    # Check for mismatches
    not_in_source = excel_folders - source_folders
    not_in_excel = source_folders - excel_folders

    if not_in_source:
        logging.warning(
            f"Folders in Excel but not in source: {', '.join(not_in_source)}")
        print(
            f"Warning: The following folders are in the Excel file but not in the source directory: {', '.join(not_in_source)}")

    if not_in_excel:
        logging.warning(
            f"Folders in source but not in Excel: {', '.join(not_in_excel)}")
        print(
            f"Warning: The following folders are in the source directory but not in the Excel file: {', '.join(not_in_excel)}")

    # Rename and move folders
    print("Starting folder renaming operation...")
    success_count = 0

    for _, row in df.iterrows():
        old_name = row['Folder Lama']
        new_name = row['Folder Baru']

        old_path = source_dir / old_name
        new_path = target_dir / new_name

        if old_path.exists() and old_path.is_dir():
            try:
                shutil.copytree(old_path, new_path)
                success_count += 1
                logging.info(f"Successfully renamed {old_name} to {new_name}")
            except Exception as e:
                logging.error(f"Error renaming folder {old_name}: {e}")
                print(f"Error renaming folder {old_name}: {e}")
        else:
            logging.warning(f"Source folder {old_name} not found")
            print(f"Warning: Source folder {old_name} not found.")

    logging.info(
        f"Operation completed. Successfully renamed {success_count} folders")
    print(
        f"Operation completed. Successfully renamed {success_count} folders.")


if __name__ == "__main__":
    try:
        rename_folders()
    except Exception as e:
        logging.exception(f"An unexpected error occurred: {e}")
        print(f"An unexpected error occurred. Please check the log file for details.")
    finally:
        input("Press Enter to exit...")
