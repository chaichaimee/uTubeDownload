# uTubeSnapshot.py

import os
import re
import glob
import subprocess
import winsound
import threading
import ui
import wx
import time
import urllib.request
import tempfile
import shutil
import json
from .uTubeDownload_core import YouTubeEXE, log, getINI, PlayWave, AddOnPath, sectionName, ConverterEXE

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

def _play_snapshot_sound():
    try:
        sound_file = os.path.join(AddOnPath, "sounds", "snapshot.wav")
        if os.path.exists(sound_file) and getINI("BeepWhileConverting"):
            winsound.PlaySound(sound_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except Exception as e:
        log(f"Error playing snapshot sound: {e}")

def _play_complete_sound():
    try:
        sound_file = os.path.join(AddOnPath, "sounds", "snapshot.wav")
        if os.path.exists(sound_file) and getINI("BeepWhileConverting"):
            winsound.PlaySound(sound_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except Exception as e:
        log(f"Error playing complete sound: {e}")

def _get_maxres_thumbnail_url(video_url):
    """Get highest resolution thumbnail URL available without compression"""
    try:
        cmd = [
            YouTubeEXE,
            video_url,
            "--skip-download",
            "--dump-json",
            "--no-playlist",
            "--no-check-certificate",
            "--quiet"
        ]
        
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        try:
            stdout, stderr = process.communicate(timeout=20)
            if process.returncode != 0:
                log(f"JSON dump error: {stderr.strip()}")
                return None
            
            # Parse JSON to find the highest resolution thumbnail
            video_info = json.loads(stdout)
            thumbnails = video_info.get('thumbnails', [])
            
            if not thumbnails:
                return None
            
            # Find maxresdefault or highest resolution
            max_res = 0
            best_url = None
            
            for thumb in thumbnails:
                # Prefer maxresdefault if available
                if thumb.get('id') == 'maxresdefault':
                    return thumb['url']
                
                # Otherwise find highest resolution
                res = thumb.get('width', 0) * thumb.get('height', 0)
                if res > max_res:
                    max_res = res
                    best_url = thumb['url']
            
            return best_url
        except subprocess.TimeoutExpired:
            process.kill()
            log("Thumbnail URL retrieval timed out")
            return None
    except Exception as e:
        log(f"Error getting maxres thumbnail URL: {str(e)}")
        return None

def _download_fullsize_thumbnail(url, save_path):
    """Download full-size thumbnail without compression"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            "Accept": "image/webp,image/apng,image/*,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Connection": "keep-alive",
            "Referer": "https://www.youtube.com/",
            "Accept-Encoding": "identity"  # Disable compression
        }
        req = urllib.request.Request(url, headers=headers)
        
        with urllib.request.urlopen(req, timeout=20) as response:
            if response.status == 200:
                # Read Content-Length to verify size
                content_length = response.headers.get('Content-Length')
                if content_length and int(content_length) < 30000:  # 30KB is too small for maxres
                    log(f"Suspicious small file size: {content_length} bytes")
                    return False
                
                with open(save_path, 'wb') as f:
                    f.write(response.read())
                
                # Verify downloaded file size
                file_size = os.path.getsize(save_path)
                if file_size < 50000:  # 50KB is too small for maxres
                    log(f"Downloaded file too small: {file_size} bytes")
                    os.remove(save_path)
                    return False
                
                return True
            else:
                log(f"HTTP error: {response.status}")
                return False
    except Exception as e:
        log(f"Full-size thumbnail download error: {str(e)}")
        return False

def _download_and_preserve_thumbnail(video_url, save_path):
    """Download and preserve original thumbnail quality"""
    temp_dir = tempfile.mkdtemp()
    try:
        temp_output = os.path.join(temp_dir, "thumbnail")
        
        # Step 1: Download original thumbnail using yt-dlp
        cmd_download = [
            YouTubeEXE,
            video_url,
            "--skip-download",
            "--write-thumbnail",
            "--no-playlist",
            "--no-check-certificate",
            "--no-part",
            "--write-all-thumbnails",  # Get all available thumbnails
            "-o", f"{temp_output}.%(ext)s"
        ]
        
        download_process = subprocess.run(
            cmd_download,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
            timeout=45
        )
        
        if download_process.returncode != 0:
            log(f"Thumbnail download failed: {download_process.stderr}")
            return False
        
        # Find downloaded thumbnail files
        thumbnail_files = glob.glob(f"{temp_output}.*")
        if not thumbnail_files:
            log("No thumbnail files found after download")
            return False
        
        # Select largest file (highest quality)
        largest_file = max(thumbnail_files, key=os.path.getsize)
        file_size = os.path.getsize(largest_file)
        
        # Verify file size is adequate (at least 50KB)
        if file_size < 50000:
            log(f"Thumbnail file too small: {file_size} bytes")
            return False
        
        # Copy to destination
        shutil.copy(largest_file, save_path)
        
        # If not JPG, convert without quality loss
        ext = os.path.splitext(largest_file)[1].lower()
        if ext != '.jpg':
            temp_jpg = os.path.join(temp_dir, "converted.jpg")
            cmd_convert = [
                ConverterEXE,
                "-i", largest_file,
                "-q:v", "1",  # Highest quality (1-31, 1 is best)
                "-y", temp_jpg
            ]
            
            convert_process = subprocess.run(
                cmd_convert,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                timeout=30
            )
            
            if convert_process.returncode != 0:
                log(f"Thumbnail conversion failed: {convert_process.stderr}")
                return False
            
            shutil.copy(temp_jpg, save_path)
        
        return os.path.exists(save_path)
    except Exception as e:
        log(f"Full-size thumbnail preservation error: {str(e)}")
        return False
    finally:
        # Cleanup temp files
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass

def capture_snapshot(video_url, download_path):
    """Capture full-size YouTube snapshot without compression"""
    if not os.path.exists(download_path):
        try:
            os.makedirs(download_path, exist_ok=True)
        except Exception as e:
            log(f"Error creating directory: {e}")
            ui.message("Error creating download folder")
            return

    next_number = _find_next_snapshot_number(download_path)
    output_filename = f"Snapshot {next_number}"
    jpg_path = os.path.join(download_path, output_filename + ".jpg")
    
    if os.path.exists(jpg_path):
        ui.message("Snapshot file already exists")
        return
    
    _play_snapshot_sound()
    ui.message("Capturing full-size snapshot...")

    def snapshot_worker():
        success = False
        file_size = 0
        
        # Method 1: Get maxres thumbnail URL
        wx.CallAfter(ui.message, "Retrieving maxres thumbnail URL...")
        thumbnail_url = _get_maxres_thumbnail_url(video_url)
        
        if thumbnail_url:
            wx.CallAfter(ui.message, "Downloading maxres thumbnail...")
            if _download_fullsize_thumbnail(thumbnail_url, jpg_path):
                file_size = os.path.getsize(jpg_path)
                if file_size >= 100000:  # At least 100KB
                    success = True
                    wx.CallAfter(ui.message, f"Maxres thumbnail captured ({file_size//1024} KB)")
        
        # Method 2: Download and preserve all thumbnails
        if not success:
            wx.CallAfter(ui.message, "Downloading original quality thumbnail...")
            if _download_and_preserve_thumbnail(video_url, jpg_path):
                file_size = os.path.getsize(jpg_path)
                if file_size >= 50000:  # At least 50KB
                    success = True
                    wx.CallAfter(ui.message, f"Original quality thumbnail captured ({file_size//1024} KB)")
        
        # Final result handling
        if success:
            wx.CallAfter(ui.message, "Full-size snapshot complete")
            wx.CallAfter(_play_complete_sound)
        else:
            wx.CallAfter(ui.message, "Error: Failed to capture full-size snapshot")
            log("All snapshot capture methods failed")
            
            # Cleanup if partial file exists
            if os.path.exists(jpg_path) and os.path.getsize(jpg_path) < 50000:
                try:
                    os.remove(jpg_path)
                except Exception:
                    pass

    # Start the snapshot worker in a new thread
    threading.Thread(target=snapshot_worker, daemon=True).start()
