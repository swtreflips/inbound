import os
import glob
from datetime import datetime

def get_latest_folders(base_dir, prefixes, limit=15):
    data = {}

    # Step 1: Collect valid date folders
    dated_folders = []
    for folder in os.listdir(base_dir):
        folder_path = os.path.join(base_dir, folder)
        if os.path.isdir(folder_path):
            try:
                folder_date = datetime.strptime(folder, "%m.%d.%y").date()
                dated_folders.append((folder_date, folder_path))
            except ValueError:
                continue

    # Step 2: Sort folders by date (descending) and take the most recent ones
    latest_folders = sorted(dated_folders, reverse=True)[:limit]

    # Step 3: Search for files in those folders
    for folder_date, folder_path in latest_folders:
        for prefix in prefixes:
            glob_pattern = os.path.join(folder_path, f"{prefix}*")
            matched_files = glob.glob(glob_pattern)
            if matched_files and prefix not in data:
                data[prefix] = os.path.normpath(folder_path).replace("\\", "/")

    return data
