import os
import sys
import tempfile
import zipfile
import subprocess
from pathlib import Path
from typing import Optional, Dict, Any, Callable
import requests

from app.version import APP_VERSION, GITHUB_REPO, is_newer_version

class AppUpdater:
    def __init__(self, repo: str = GITHUB_REPO, current_version: str = APP_VERSION):
        self.repo = repo
        self.current_version = current_version
        self.headers = {
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": f"AssignmentFixer-Updater/{current_version}"
        }

    def check_for_update(self) -> Optional[Dict[str, Any]]:
        """
        Query GitHub Releases API for the latest release.
        Returns release info dict if a newer version is available, else None.
        """
        url = f"https://api.github.com/repos/{self.repo}/releases/latest"
        try:
            resp = requests.get(url, headers=self.headers, timeout=10)
            if resp.status_code == 404:
                return None
            resp.raise_for_status()
            data = resp.json()

            tag_name = data.get("tag_name", "")
            if not is_newer_version(tag_name, self.current_version):
                return None

            # Find matching zip asset
            assets = data.get("assets", [])
            selected_asset = None
            for asset in assets:
                name = asset.get("name", "").lower()
                if name.endswith(".zip"):
                    # Prioritize asset with AssignmentFixer or windows in name
                    if "assignmentfixer" in name or "windows" in name:
                        selected_asset = asset
                        break
                    elif selected_asset is None:
                        selected_asset = asset

            if not selected_asset:
                # Fallback to source zipball if no release asset attached
                download_url = data.get("zipball_url", "")
                asset_name = f"AssignmentFixer-{tag_name}.zip"
                asset_size = 0
            else:
                download_url = selected_asset.get("browser_download_url", "")
                asset_name = selected_asset.get("name", "")
                asset_size = selected_asset.get("size", 0)

            return {
                "tag_name": tag_name,
                "name": data.get("name") or tag_name,
                "body": data.get("body", "").strip(),
                "published_at": data.get("published_at", ""),
                "html_url": data.get("html_url", ""),
                "download_url": download_url,
                "asset_name": asset_name,
                "asset_size": asset_size,
            }
        except Exception as e:
            print(f"[Updater] Error checking update: {e}")
            return None

    def download_update(
        self,
        download_url: str,
        dest_path: str,
        progress_cb: Optional[Callable[[int, int], None]] = None
    ) -> bool:
        """
        Stream download update package with progress callback (bytes_downloaded, total_bytes).
        """
        try:
            with requests.get(download_url, headers=self.headers, stream=True, timeout=30) as r:
                r.raise_for_status()
                total_bytes = int(r.headers.get("content-length", 0))
                downloaded = 0
                os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                with open(dest_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=65536):
                        if chunk:
                            f.write(chunk)
                            downloaded += len(chunk)
                            if progress_cb:
                                progress_cb(downloaded, total_bytes)
            return True
        except Exception as e:
            print(f"[Updater] Download failed: {e}")
            return False

    def launch_update_script(self, zip_path: str, target_dir: Optional[Path] = None) -> bool:
        """
        Extracts update zip and creates a detached Windows batch script that waits for
        the current application process to close, overwrites files, and restarts the app.
        """
        is_frozen = getattr(sys, "frozen", False)
        if target_dir is None:
            if is_frozen:
                target_dir = Path(sys.executable).resolve().parent
            else:
                target_dir = Path(__file__).resolve().parent.parent.parent / "dist" / "AssignmentFixer"
                if not target_dir.exists():
                    target_dir = Path(__file__).resolve().parent.parent.parent

        extract_dir = Path(tempfile.mkdtemp(prefix="asfixer_upd_"))
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(extract_dir)
        except Exception as e:
            print(f"[Updater] Failed to extract zip: {e}")
            return False

        # If the zip extracted to a single nested folder, point to that folder
        subdirs = [p for p in extract_dir.iterdir() if p.is_dir()]
        subfiles = [p for p in extract_dir.iterdir() if p.is_file()]
        source_dir = extract_dir
        if len(subdirs) == 1 and len(subfiles) == 0:
            source_dir = subdirs[0]

        # Determine relaunch command
        current_pid = os.getpid()
        if is_frozen:
            relaunch_target = str(Path(sys.executable).resolve())
        else:
            exe_cand = target_dir / "AssignmentFixer.exe"
            if exe_cand.exists():
                relaunch_target = str(exe_cand)
            else:
                relaunch_target = str(Path(sys.executable).resolve())

        # Write updater batch file
        batch_path = extract_dir.parent / f"apply_update_{current_pid}.bat"
        batch_content = f"""@echo off
chcp 65001 >nul
echo [AssignmentFixer Updater] Dang cho dong ung dung (PID {current_pid})...
timeout /t 2 /nobreak >nul

:WAIT_PID
tasklist /FI "PID eq {current_pid}" 2>NUL | find /I "{current_pid}" >NUL
if "%ERRORLEVEL%"=="0" (
    timeout /t 1 /nobreak >nul
    goto WAIT_PID
)

echo [AssignmentFixer Updater] Dang sao chep tep cap nhat vao {target_dir}...
xcopy /s /e /y /q "{source_dir}\\*" "{target_dir}\\"

echo [AssignmentFixer Updater] Dang don dep tap tin tam...
rd /s /q "{extract_dir}" 2>nul
del /f /q "{zip_path}" 2>nul

echo [AssignmentFixer Updater] Dang khoi dong lai ung dung...
start "" "{relaunch_target}"

del /f /q "%~f0" 2>nul
exit
"""
        with open(batch_path, "w", encoding="utf-8") as bf:
            bf.write(batch_content)

        # Launch detached batch script
        try:
            subprocess.Popen(
                ["cmd.exe", "/c", str(batch_path)],
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
                close_fds=True
            )
            return True
        except Exception as e:
            print(f"[Updater] Failed to start update batch script: {e}")
            return False
