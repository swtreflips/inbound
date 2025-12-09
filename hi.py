import os

def get_onedrive_business_path():
    """Return the OneDrive Business root folder dynamically."""
    
    # 1. Try environment variables first
    for key in ("OneDriveCommercial", "OneDriveBusiness", "OneDrive"):
        path = os.environ.get(key)
        if path and os.path.exists(path):
            return path
    
    # 2. Search user directory as fallback
    user_dir = os.path.expanduser("~")
    for name in os.listdir(user_dir):
        if "OneDrive" in name and os.path.isdir(os.path.join(user_dir, name)):
            # Only pick business tenants: require a hyphen " - "
            if " - " in name:
                return os.path.join(user_dir, name)
    
    raise FileNotFoundError("Could not detect OneDrive Business folder.")

# ------- USE IT HERE --------

one_drive_root = get_onedrive_business_path()
base_dir = os.path.join(one_drive_root, "Inbound Update")

print("Resolved OneDrive path:", one_drive_root)
print("Using base_dir:", base_dir)
