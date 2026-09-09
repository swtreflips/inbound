"""Locate the shared "Inbound Update" folder on any machine.

The folder reaches different people in different shapes:

  1. Owner                       C:\\Users\\<user>\\OneDrive - Prime Time Packaging\\Inbound Update
  2. Colleague, "Add shortcut    C:\\Users\\<user>\\OneDrive - Prime Time Packaging\\Inbound Update
     to My files"               (same shape - the shortcut lands in their OneDrive root)
  3. Colleague syncing the       C:\\Users\\<user>\\Prime Time Packaging\\<Library>\\Inbound Update
     SharePoint library          (no "OneDrive - " prefix at all)

Case 3 is why hardcoding "~/OneDrive - Prime Time Packaging" is not enough. The OneDrive
client records every sync root in the registry under
HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\<Account>, where `UserFolder` is the OneDrive
root and the *value names* under `Tenants\\<Tenant>` are the local mount paths of synced
libraries. We read those, then look for the target folder beneath each.

Usage:
    from onedrive_path import get_inbound_base_dir
    base_dir = get_inbound_base_dir()

Escape hatch - if someone's layout is unusual, set an environment variable and we use it
verbatim, no searching:
    setx INBOUND_UPDATE_DIR "D:\\some\\path\\Inbound Update"
"""

import os
import re
from typing import List, Optional

try:
    import winreg  # Windows only; absence just disables the registry probe
except ImportError:  # pragma: no cover
    winreg = None

TARGET_NAME = "Inbound Update"
ENV_OVERRIDE = "INBOUND_UPDATE_DIR"
MAX_DEPTH = 3

# The script's own convention: dated drop folders named MM.DD.YY
_DATE_FOLDER = re.compile(r"^\d{2}\.\d{2}\.\d{2}$")


def _registry_roots() -> List[str]:
    """Every OneDrive sync root this user has: personal/business roots and synced libraries."""
    roots: List[str] = []
    if winreg is None:
        return roots
    try:
        accounts_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                      r"Software\Microsoft\OneDrive\Accounts")
    except OSError:
        return roots

    with accounts_key:
        for i in range(1024):
            try:
                account = winreg.EnumKey(accounts_key, i)
            except OSError:
                break
            try:
                with winreg.OpenKey(accounts_key, account) as acct:
                    # the account's own root
                    try:
                        folder, _ = winreg.QueryValueEx(acct, "UserFolder")
                        if folder:
                            roots.append(folder)
                    except OSError:
                        pass
                    # synced SharePoint/Teams libraries: value NAMES are local paths
                    try:
                        with winreg.OpenKey(acct, "Tenants") as tenants:
                            for j in range(1024):
                                try:
                                    tenant = winreg.EnumKey(tenants, j)
                                except OSError:
                                    break
                                try:
                                    with winreg.OpenKey(tenants, tenant) as t:
                                        for n in range(4096):
                                            try:
                                                name, _val, _typ = winreg.EnumValue(t, n)
                                            except OSError:
                                                break
                                            if name and (":" in name or name.startswith("\\\\")):
                                                roots.append(name)
                                except OSError:
                                    continue
                    except OSError:
                        pass
            except OSError:
                continue
    return roots


def _env_roots() -> List[str]:
    return [os.environ[k] for k in ("OneDriveCommercial", "OneDriveBusiness", "OneDrive")
            if os.environ.get(k)]


def _profile_roots() -> List[str]:
    """Last resort: top-level folders in the user profile that look like tenant sync roots."""
    found: List[str] = []
    home = os.path.expanduser("~")
    try:
        entries = os.listdir(home)
    except OSError:
        return found
    for name in entries:
        path = os.path.join(home, name)
        if not os.path.isdir(path):
            continue
        # "OneDrive - Tenant" (business) or a bare tenant folder from a synced library
        if name.startswith("OneDrive") or " - " in name:
            found.append(path)
    return found


def _candidate_roots() -> List[str]:
    seen = set()
    ordered: List[str] = []
    for root in _registry_roots() + _env_roots() + _profile_roots():
        if not root:
            continue
        key = os.path.normcase(os.path.normpath(root))
        if key in seen:
            continue
        seen.add(key)
        if os.path.isdir(root):
            ordered.append(os.path.normpath(root))
    return ordered


def _find_under(root: str, target: str, max_depth: int = MAX_DEPTH) -> List[str]:
    """Find directories named `target` beneath `root`, up to `max_depth` levels."""
    # Fast path: the overwhelmingly common case is target sitting directly in the root.
    direct = os.path.join(root, target)
    if os.path.isdir(direct):
        return [direct]

    hits: List[str] = []
    base = root.rstrip(os.sep).count(os.sep)
    for dirpath, dirnames, _files in os.walk(root):
        if dirpath.count(os.sep) - base >= max_depth:
            dirnames[:] = []
            continue
        # don't descend into dot-dirs, Office lock files, or node_modules-style noise
        dirnames[:] = [d for d in dirnames if not d.startswith((".", "~$"))]
        for d in dirnames:
            if d.lower() == target.lower():
                hits.append(os.path.join(dirpath, d))
    return hits


def _looks_like_inbound(path: str) -> bool:
    """True if the folder contains MM.DD.YY dated drop folders - confirms it's the right one."""
    try:
        for name in os.listdir(path):
            if _DATE_FOLDER.match(name) and os.path.isdir(os.path.join(path, name)):
                return True
    except OSError:
        pass
    return False


def get_inbound_base_dir(target: str = TARGET_NAME, verbose: bool = False) -> str:
    """Return the absolute path to the shared folder. Raises FileNotFoundError if not found."""
    override = os.environ.get(ENV_OVERRIDE)
    if override:
        if os.path.isdir(override):
            if verbose:
                print(f"[onedrive_path] using {ENV_OVERRIDE}={override}")
            return os.path.normpath(override)
        raise FileNotFoundError(
            f"{ENV_OVERRIDE} is set to '{override}' but that folder does not exist."
        )

    roots = _candidate_roots()
    if verbose:
        print("[onedrive_path] candidate sync roots:")
        for r in roots:
            print(f"    {r}")

    matches: List[str] = []
    seen = set()
    for root in roots:
        for hit in _find_under(root, target):
            key = os.path.normcase(os.path.normpath(hit))
            if key not in seen:
                seen.add(key)
                matches.append(os.path.normpath(hit))

    if not matches:
        raise FileNotFoundError(
            f"Could not find a '{target}' folder in any OneDrive sync root.\n"
            f"Searched: {', '.join(roots) if roots else '(no sync roots detected)'}\n\n"
            f"Fixes:\n"
            f"  - In OneDrive on the web, open the shared folder and choose\n"
            f"    'Add shortcut to My files', then let it sync.\n"
            f"  - Or point the script at it directly:\n"
            f'      setx {ENV_OVERRIDE} "C:\\path\\to\\{target}"'
        )

    # Prefer a folder that actually contains dated drop folders.
    confirmed = [m for m in matches if _looks_like_inbound(m)]
    chosen = (confirmed or matches)[0]

    if verbose or len(matches) > 1:
        if len(matches) > 1:
            print(f"[onedrive_path] multiple '{target}' folders found; using: {chosen}")
            for m in matches:
                if m != chosen:
                    print(f"    (also saw) {m}")
        else:
            print(f"[onedrive_path] resolved: {chosen}")

    if not confirmed:
        print(f"[onedrive_path] warning: '{chosen}' has no MM.DD.YY drop folders yet - "
              f"it may still be syncing.")

    return chosen


if __name__ == "__main__":
    print(get_inbound_base_dir(verbose=True))
