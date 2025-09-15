# uTubeSnapshot.py

import os
import re
import glob
import subprocess
import threading
import ui
import wx
import time
import shutil
import addonHandler
from .uTubeDownload_core import YouTubeEXE, log, getINI, PlayWave, AddOnPath, sectionName

addonHandler.initTranslation()

def _find_next_snapshot_number(save_path):
    try:
        existing_files = glob.glob(os.path.join(save_path, "Snapshot *.jpg"))
        numbers = []
        for file_path in existing_files:
            match = re.search(r"Snapshot (\d+)\.jpg$", os.path.basename(file_path))
            if match:
                numbers.append(int(match.group(1)))
        if not numbers:
            return 1
        next_number = max(numbers) + 1
        return next_number
    except Exception as e:
        log(f"Error finding next snapshot number: {e}")
        return 1

def capture_snapshot(video_url, download_path):
    """Capture full-size YouTube snapshot using yt-dlp to find and download the best thumbnail."""
    if not os.path.exists(download_path):
        try:
            os.makedirs(download_path, exist_ok=True)
        except Exception as e:
            log(f"Error creating directory: {e}")
            wx.CallAfter(ui.message, _("Error creating download folder"))
            return

    next_number = _find_next_snapshot_number(download_path)
    output_filename = f"Snapshot {next_number}"
    
    temp_dir = os.path.join(download_path, "temp_snapshot_dir")
    os.makedirs(temp_dir, exist_ok=True)
    temp_output_path = os.path.join(temp_dir, f"{output_filename}.%(ext)s")
    
    final_output_path = os.path.join(download_path, f"{output_filename}.jpg")

    if os.path.exists(final_output_path):
        wx.CallAfter(ui.message, _("Snapshot file already exists"))
        return

    PlayWave("snapshot")
    wx.CallAfter(ui.message, _("Capturing full-size snapshot..."))

    def snapshot_worker():
        success = False
        try:
            wx.CallAfter(ui.message, _("Downloading thumbnail..."))

            cmd = [
                YouTubeEXE,
                video_url,
                "--write-thumbnail",
                "--skip-download",
                "--no-playlist",
                "--no-check-certificate",
                "--convert-thumbnails", "jpg",
                "-o", temp_output_path
            ]
            
            process = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            wx.CallAfter(ui.message, _("Processing snapshot..."))

            if process.returncode != 0:
                log(f"Snapshot capture failed: {process.stderr}")
                wx.CallAfter(ui.message, _("Error: Failed to capture snapshot."))
                PlayWave("error")
                return

            downloaded_files = glob.glob(os.path.join(temp_dir, f"{output_filename}*.jpg"))
            if not downloaded_files:
                log("No JPEG thumbnail file found after download")
                wx.CallAfter(ui.message, _("Error: No snapshot file created."))
                PlayWave("error")
                return
            
            downloaded_file = downloaded_files[0]
            
            # Remove file size check to allow small thumbnails
            shutil.move(downloaded_file, final_output_path)
            success = True
            
        except Exception as e:
            log(f"An unexpected error occurred during snapshot capture: {e}")
            wx.CallAfter(ui.message, _("An unexpected error occurred."))
            PlayWave("error")
        finally:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            
            if success:
                wx.CallAfter(ui.message, _("Full-size snapshot complete"))
                PlayWave("complete")
            else:
                PlayWave("error")
    
    threading.Thread(target=snapshot_worker, daemon=True).start()

