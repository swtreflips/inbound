import time
import glob
from typing import Optional
import pandas as pd
from typing import Dict
import xlwings as xw
from datetime import datetime
import os
from find_latest_folders import get_latest_folders
from tqdm import tqdm

# Functions
ERP_df = None
Topocean_df = None
OEC_Portal_df = None
OEC_Email_df = None
Soma_df = None
TaneraGo_df = None
Harbour_df = None
Idc_df = None

# File Pattern Recognition
def load_file_from_config(config, folder_path):
    # Handle both "extensions" and "extension"/"loader" format
    if "extensions" in config:
        ext_loader_pairs = config["extensions"]
    elif "extension" in config and "loader" in config:
        extensions = config["extension"]
        if isinstance(extensions, str):
            extensions = [extensions]
        ext_loader_pairs = [(ext, config["loader"]) for ext in extensions]
    else:
        raise KeyError("Missing required keys: 'extensions' or 'extension' with 'loader'")

    for ext, loader in ext_loader_pairs:
        pattern = os.path.join(folder_path, f"{config['prefix']}*.{ext}")
        matching_files = glob.glob(pattern)
        if matching_files:
            file_path = matching_files[0]
            try:
                df = loader(file_path, **config.get("kwargs", {}))
                if "postprocess" in config:
                    df = config["postprocess"](df)
                return df
            except Exception as e:
                print(f"Error loading {file_path}: {e}")
                return None
    return None
    
# OEC Portal Data Cleaning    
def clean_oec_portal(df):
    # Convert columns 8 to 10 (i.e., columns at index 8, 9, 10) to datetime.date
    for col in df.columns[8:11]:
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

    # Create new column: expected delivery date = ETA-Last CY/CFS Location + 5 days
    df['expected delivery date'] = pd.to_datetime(df['ETA-Last CY/CFS Location'], errors='coerce') + pd.Timedelta(days=5)

    # Move 'expected delivery date' to column index 10
    df.insert(10, 'expected delivery date', df.pop('expected delivery date'))

    return df

# Soma Data Cleaning
def clean_soma(df: pd.DataFrame) -> pd.DataFrame:
    # Drop unwanted columns
    df = df.drop(columns=["SCAC CODE", "AN STATUS"])

    # Fix Master Bill of Lading
    prefix_to_scac = {
        'BOM': 'HDMU',
        'MUM': 'ONEY',
        'BO': 'HLCU',
        '067': 'WHLC',
        '639': 'COSU',
    }
    sorted_prefixes = sorted(prefix_to_scac.keys(), key=lambda x: len(x), reverse=True)

    def update_mbl(row):
        for prefix in sorted_prefixes:
            if isinstance(row, str) and row.startswith(prefix):
                return prefix_to_scac[prefix] + row
        return row

    df['Master Bill of Lading'] = df['Master Bill of Lading'].apply(update_mbl)

    # Format specific columns as dates
    for col in df.columns[8:15]:
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

    return df

# Tanera Go Data Cleaning
def clean_tanerago(df: pd.DataFrame) -> pd.DataFrame:
    # Format specific columns as dates
    for col in df.columns[15:29]:
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
    return df

# Opening Template and Pasting Dataframes
def load_template_and_paste_data(
    prefix: str = "Inbound Weekly Update Template",
    sheets_data: Dict[str, pd.DataFrame] = None
) -> None:
    file_pattern = f"{prefix}*.xlsm"
    matching_files = glob.glob(file_pattern)

    if not matching_files:
        print(f"No matching files found for prefix: {prefix}")
        return

    file_to_load = matching_files[0]
    app = xw.App(visible=False)
    wb = app.books.open(file_to_load)
    print(f"Loaded workbook: {file_to_load}")

    # Close extra open workbooks
    for book in app.books:
        if book.name != wb.name:
            print(f"Closing extra workbook: {book.name}")
            book.close()

    # Paste each DataFrame into its corresponding sheet at A2
# Paste each DataFrame into its corresponding sheet at A2
    if sheets_data:
        for sheet_name, df in sheets_data.items():
            if df is None:
                print(f"Warning: DataFrame for '{sheet_name}' is None, skipping.")
                continue
            try:
                sheet = wb.sheets[sheet_name]
                sheet.range("A2").value = df.values
                print(f"Pasted data into sheet: {sheet_name}")
            except Exception as e:
                print(f"Error pasting into sheet '{sheet_name}': {e}")

    # Save workbook into today's folder under ~/Documents/Reports/
    today = datetime.today()
    today_str = today.strftime("%m.%d.%y")
    folder_name = today.strftime("%m.%d.%y")
    reports_dir = os.path.expanduser(f"~/OneDrive - Prime Time Packaging/Inbound Update")

    # Create the folder if it doesn't exist
    os.makedirs(reports_dir, exist_ok=True)

    new_filename = f"Inbound Weekly Update Template LMH {today_str}.xlsm"
    new_filepath = os.path.join(reports_dir, new_filename)
    wb.save(new_filepath)
    print(f"Workbook saved as: {new_filename} in {reports_dir}")

    # Close workbook and quit Excel
    wb.close()
    print("Workbook closed.")
    app.quit()
    print("Excel app quit.")


# File Configs and variables

file_configs = {
    "ERP": {
        "prefix": "InboundShipments",
        "extension": "csv",
        "loader": pd.read_csv,
        "kwargs": {}
    },
    "Topocean": {
        "prefix": "PRIME TIME PACKAGING",
        "extension": "xls",
        "loader": pd.read_excel,
        "kwargs": {"engine": "xlrd", "skiprows": 7}
    },
    "OEC_Email": {
        "prefix": "OEC GROUP Container Tracking Report",
        "extension": "xlsx",
        "loader": pd.read_excel,
        "kwargs": {"skiprows": 6},
        "postprocess": lambda df: df.drop(columns=["Unnamed: 0", "Unnamed: 1"], errors="ignore")
    },
    "OEC_Portal": {
        "prefix": "OEC2 Upload",
        "extensions": [("xlsx", pd.read_excel), ("csv", pd.read_csv)],
        "loader": pd.read_excel,
        "kwargs": {},
        "postprocess": clean_oec_portal 
    },
    "Soma": {
        "prefix": "PTP SOMA",
        "extension": "xlsx",
        "loader": pd.read_excel,
        "kwargs": {"skiprows": [1]},
        "postprocess": clean_soma
    },
    "TaneraGo": {
        "prefix": "Shipment_status",
        "extension": "xlsx",
        "loader": pd.read_excel,
        "kwargs": {},
        "postprocess": clean_tanerago
    },
    "Harbour": {
        "prefix": "Forwarder Inbound Template",
        "extension": "xlsx",
        "loader": pd.read_excel,
        "kwargs": {}
    },
    "Idc": {
        "prefix": "PRIME TIME DSR",
        "extensions": [("xlsx", pd.read_excel), ("csv", pd.read_csv)],
        "loader": pd.read_excel,
        "kwargs": {}
    }
}

# Define base directory
base_dir = os.path.expanduser("~/OneDrive - Prime Time Packaging/Inbound Update")
prefixes = [cfg["prefix"] for cfg in file_configs.values()]
latest_folders = get_latest_folders(base_dir, prefixes)

# Data Cleaning Code Execution
start_time = time.time()
# Data Cleaning Code Execution
for config_name, config in file_configs.items():
    prefix = config["prefix"]
    folder_path = latest_folders.get(prefix)

    if not folder_path:
        print(f"No folder found for prefix '{prefix}'")
        continue

    df = load_file_from_config(config, folder_path)

    if df is not None:
        globals()[f"{config_name}_df"] = df
        print(f"Loaded {config_name} file successfully!")
    else:
        print(f"Failed to load {config_name} file.")



# Dictionary mapping each DataFrame to its corresponding sheet name
sheets_data = {
    'ERP': ERP_df,
    'topocean': Topocean_df,
    'OEC Portal': OEC_Portal_df,
    'OEC Email': OEC_Email_df,
    'Soma': Soma_df,
    'Tanera Go': TaneraGo_df,
    'Harbour': Harbour_df,
    'IDC': Idc_df  
}

# Call the function 
load_template_and_paste_data(sheets_data=sheets_data)

# <<< Place timing code here, after everything else >>>
end_time = time.time()
elapsed_time = end_time - start_time
print(f"\nScript executed in {elapsed_time:.2f} seconds.")