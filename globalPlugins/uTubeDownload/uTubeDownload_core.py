# uTubeDownload_core.py

import wx
import os
import json
import time
import urllib
import threading
import subprocess
import datetime
import glob
import winsound
import api
import controlTypes
import speech
import ui
import config
from scriptHandler import script
import gui
import re
import uuid
import tones
import shutil
import tempfile
import psutil
import sys
from queue import Queue

AddOnSummary = "uTubeDownload"
AddOnName = "uTubeDownload"
if sys.version_info.major >= 3 and sys.version_info.minor >= 10:
    AddOnPath = os.path.dirname(__file__)
else:
    AddOnPath = os.path.dirname(__file__)
ToolsPath = os.path.join(AddOnPath, "Tools")
SoundPath = os.path.join(AddOnPath, "sounds")
AppData = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming')
DownloadPath = None
sectionName = AddOnName
_download_queue = Queue()
_heartbeat_thread = None
_heartbeat_active = False
Aria2cEXE = os.path.join(ToolsPath, "aria2c.exe")
YouTubeEXE = os.path.join(ToolsPath, "yt-dlp.exe")
ConverterEXE = os.path.join(ToolsPath, "ffmpeg.exe")
ConverterPath = ToolsPath
_global_state_lock = threading.Lock()
_global_active_downloads = 0
_global_active_lock = threading.Lock()
_num_workers = config.conf[sectionName]["MaxConcurrentDownloads"]

def getStateFilePath():
    try:
        import globalVars
        if globalVars.appArgs.secure:
            return os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'nvda', 'uTubeDownload.json')
        configDir = globalVars.appArgs.configPath or os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'nvda')
        return os.path.join(configDir, 'uTubeDownload.json')
    except Exception:
        return os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'nvda', 'uTubeDownload.json')
StateFilePath = getStateFilePath()

def getINI(key):
    return config.conf[sectionName][key]

def setINI(key, value):
    config.conf[sectionName][key] = value

def PlayWave(filename):
    try:
        path = os.path.join(SoundPath, filename + ".wav")
        if os.path.exists(path) and getINI("BeepWhileConverting"):
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except Exception as e:
        log(f"Error playing sound: {e}")

def _heartbeat_loop():
    global _heartbeat_active
    while _heartbeat_active:
        PlayWave("heart")
        time.sleep(4)
    winsound.PlaySound(None, winsound.SND_PURGE)

def startHeartbeat():
    global _heartbeat_thread, _heartbeat_active
    if not _heartbeat_active:
        _heartbeat_active = True
        _heartbeat_thread = threading.Thread(target=_heartbeat_loop, daemon=True)
        _heartbeat_thread.start()

def stopHeartbeat():
    global _heartbeat_thread, _heartbeat_active
    _heartbeat_active = False
    if _heartbeat_thread and _heartbeat_thread.is_alive():
        try:
            _heartbeat_thread.join()
        except Exception:
            pass

def initialize_folders():
    global DownloadPath
    folder = getINI("ResultFolder") or os.path.join(AppData, "uTubeDownload")
    setINI("ResultFolder", folder)
    DownloadPath = folder
    if not os.path.exists(DownloadPath):
        os.makedirs(DownloadPath, exist_ok=True)
    if not os.path.exists(ToolsPath):
        os.makedirs(ToolsPath, exist_ok=True)
    if not os.path.exists(SoundPath):
        os.makedirs(SoundPath, exist_ok=True)
    if not os.path.exists(StateFilePath):
        saveState([])
    log("Initialized folders")

def saveState(queue_list):
    try:
        os.makedirs(os.path.dirname(StateFilePath), exist_ok=True)
        with open(StateFilePath, 'w', encoding='utf-8') as f:
            json.dump(queue_list, f, ensure_ascii=False, indent=4)
    except Exception as e:
        log(f"Error saving state: {e}")

def loadState():
    try:
        if os.path.exists(StateFilePath):
            with open(StateFilePath, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        log(f"Error loading state: {e}")
    return []

def clearState():
    try:
        if os.path.exists(StateFilePath):
            with open(StateFilePath, 'w', encoding='utf-8') as f:
                json.dump([], f)
    except Exception as e:
        log(f"Error clearing state: {e}")

def addDownloadToQueue(download_obj):
    with _global_state_lock:
        queue = loadState()
        download_obj["id"] = str(uuid.uuid4())
        download_obj["start_time"] = datetime.datetime.now().isoformat()
        download_obj["status"] = "queued"
        queue.append(download_obj)
        saveState(queue)
        log(f"Added download to queue: ID {download_obj['id']}")
        return download_obj["id"]

def updateDownloadStatusInQueue(download_id, status):
    with _global_state_lock:
        queue = loadState()
        updated = False
        for item in queue:
            if item.get("id") == download_id:
                item["status"] = status
                if status in ["completed", "failed", "cancelled"]:
                    item["end_time"] = datetime.datetime.now().isoformat()
                updated = True
                break
        if updated:
            saveState(queue)
    log(f"Updated download status: ID {download_id} to {status}")

def removeCompletedOrFailedDownloadsFromQueue():
    with _global_state_lock:
        queue = loadState()
        new_queue = [item for item in queue if item.get("status") not in ["completed", "failed", "cancelled"]]
        if len(new_queue) < len(queue):
            saveState(new_queue)
            log(f"Removed {len(queue) - len(new_queue)} completed/failed downloads from queue")

def makePrintable(s):
    return "".join(c if c.isprintable() else " " for c in str(s))

def validFilename(s):
    s = str(s).strip().replace(" ", "_")
    s = re.sub(r'(?u)[^-\w.]', '', s)
    return s

def log(s):
    try:
        api.log.info(f"uTubeDownload: {makePrintable(s)}")
        if getINI("Logging"):
            path = getINI("ResultFolder") or DownloadPath
            os.makedirs(path, exist_ok=True)
            with open(os.path.join(path, "log.txt"), "a", encoding="utf-8") as f:
                f.write(f"{datetime.datetime.now()} - {makePrintable(s)}\n")
    except Exception as e:
        api.log.error(f"uTubeDownload: Error writing log: {e}")

def createFolder(folder):
    if not os.path.isdir(folder):
        try:
            os.makedirs(folder, exist_ok=True)
            log(f"Created folder: {folder}")
            return True
        except Exception as e:
            ui.message(_("Cannot create folder"))
            log(f"Failed to create folder: {e}")
            return False
    return True

def getCurrentAppName():
    try:
        return api.getForegroundObject().appModule.appName
    except Exception:
        return "Unknown"

def isBrowser():
    obj = api.getFocusObject()
    return obj.treeInterceptor is not None

def getCurrentDocumentURL():
    try:
        obj = api.getFocusObject()
        if hasattr(obj, 'treeInterceptor') and obj.treeInterceptor:
            try:
                url = obj.treeInterceptor.documentConstantIdentifier
                if url:
                    return urllib.parse.unquote(url)
            except Exception:
                pass
    except Exception as e:
        log(f"Error getting URL: {e}")
    return None

def getLinkURL():
    obj = api.getNavigatorObject()
    if obj.role == controlTypes.Role.LINK:
        url = obj.value
        if url:
            url = urllib.parse.unquote(url)
            return url[:-1] if url.endswith("/") else url
    return ""

def getLinkName():
    obj = api.getNavigatorObject()
    if obj.role == controlTypes.Role.LINK:
        return validFilename(obj.name)
    return ""

def getMultimediaURLExtension():
    url = getLinkURL()
    return url[url.rfind("."):].lower() if "." in url else ""

def isValidMultimediaExtension(ext):
    return ext.replace(".", "") in {
        "aac", "avi", "flac", "mkv", "m3u8", "m4a", "m4s", "m4v",
        "mpg", "mov", "mp2", "mp3", "mp4", "mpeg", "mpegts", "ogg",
        "ogv", "oga", "ts", "vob", "wav", "webm", "wmv", "f4v",
        "flv", "swf", "avchd", "3gp"
    }

def getWebSiteTitle():
    try:
        title = api.getForegroundObject().name
        unwanted_suffixes = [" - YouTube", "| YouTube", " - Google Chrome", " - Brave", " - Microsoft Edge"]
        for suffix in unwanted_suffixes:
            title = title.replace(suffix, "")
        return title
    except Exception:
        return "Unknown_Title"

def checkFileExists(savePath, title, extension, is_trimming=False):
    if not getINI("SkipExisting"):
        return False
    
    sanitized_title = validFilename(title)
    filename = os.path.join(savePath, f"{sanitized_title}.{extension}")
    
    if is_trimming:
        # For trimming, allow download even if file exists
        return False
    
    if os.path.exists(filename):
        log(f"File '{filename}' already exists.")
        return True
    
    temp_patterns = [
        f"{sanitized_title}.part",
        f"{sanitized_title}.ytdl",
        f"{sanitized_title}.temp",
        f"{sanitized_title}.download",
        f"{sanitized_title}.f*.tmp",
        f"{sanitized_title}.f*.webm",
        f"{sanitized_title}.f*.m4a",
        f"{sanitized_title}.f*.mp4",
        f"{sanitized_title}.part.aria2",
        f"{sanitized_title}.aria2"
    ]
    
    for pattern in temp_patterns:
        full_pattern = os.path.join(savePath, pattern)
        if glob.glob(full_pattern):
            log(f"Found temp file matching pattern {full_pattern}, not skipping.")
            return False
            
    return False

def safeMessageBox(message, title, style):
    return gui.messageBox(message, title, style)

def promptResumeDownloads(downloads_list):
    count = len(downloads_list)
    msg = _("Found {count} interrupted downloads\nResume all?").format(count=count)
    return safeMessageBox(msg, _("Resume downloads"), wx.YES_NO) == wx.YES

def _cleanup_temp_files(save_path, title, file_format, check_count=2):
    if not title or not save_path:
        log(f"Temp cleanup skipped: title or path missing (title: {title}, path: {save_path})")
        return
    sanitized_title = validFilename(title)
    base_filename = os.path.join(save_path, sanitized_title)
    
    temp_patterns = [
        f"{sanitized_title}.part",
        f"{sanitized_title}.ytdl",
        f"{sanitized_title}.temp",
        f"{sanitized_title}.download",
        f"{sanitized_title}.f*.tmp",
        f"{sanitized_title}.f*.webm",
        f"{sanitized_title}.f*.m4a",
        f"{sanitized_title}.f*.mp4",
        f"{sanitized_title}.part.aria2",
        f"{sanitized_title}.aria2"
    ]
    
    if file_format == "mp3":
        temp_patterns.append(f"{sanitized_title}.mp4")
    
    final_file = os.path.join(save_path, f"{sanitized_title}.{file_format}")
    
    for _ in range(check_count):
        for pattern in temp_patterns:
            for temp_file in glob.glob(os.path.join(save_path, pattern)):
                # Skip the final output file
                if temp_file == final_file:
                    continue
                # Only delete temporary files
                if ('f' in os.path.basename(temp_file).split('.')[0] or 
                    temp_file.endswith(('.part', '.ytdl', '.temp', '.download', '.aria2', '.part.aria2', '.mp4'))):
                    try:
                        os.remove(temp_file)
                        log(f"Removed temp file: {temp_file}")
                    except Exception as e:
                        log(f"Error removing temp file {temp_file}: {e}")

def get_video_duration(url):
    try:
        cmd = [YouTubeEXE, "--get-duration", "--no-playlist", url]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        if result.returncode == 0:
            duration_str = result.stdout.strip()
            if duration_str:
                parts = duration_str.split(':')
                if len(parts) == 3:
                    h, m, s = map(int, parts)
                    return h * 3600 + m * 60 + s
                elif len(parts) == 2:
                    m, s = map(int, parts)
                    return m * 60 + s
                elif len(parts) == 1:
                    return int(parts[0])
                else:
                    return None
        return None
    except Exception as e:
        log(f"Error getting video duration: {e}")
    return None

def get_file_duration(file_path):
    try:
        cmd = [ConverterEXE, "-i", file_path, "-show_entries", "format=duration", "-v", "quiet", "-of", "csv=p=0"]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        if result.returncode == 0:
            duration_str = result.stdout.strip()
            if duration_str:
                return float(duration_str)
        return None
    except Exception as e:
        log(f"Error getting file duration: {e}")
    return None

def repairIncompleteFiles(path):
    repaired_count = 0
    patterns = [
        "*.part", "*.ytdl", "*.temp", "*.download", "*.f*.tmp",
        "*.f*.webm", "*.f*.m4a", "*.f*.mp4", "*.part.aria2", "*.aria2"
    ]
    
    for pattern in patterns:
        full_pattern = os.path.join(path, pattern)
        for temp_file in glob.glob(full_pattern):
            try:
                base_name, _ = os.path.splitext(temp_file)
                if base_name.endswith('.part') or base_name.endswith('.aria2'):
                    base_name, _ = os.path.splitext(base_name)
                
                original_file = os.path.splitext(base_name)[0]
                
                if os.path.exists(original_file + '.mp4') or os.path.exists(original_file + '.mp3'):
                    log(f"Skipping repair for {temp_file}: corresponding file already exists.")
                    continue
                
                matches = re.findall(r"^(.*?)(?:-\w+)?(?:\.\w+)?$", original_file)
                if matches:
                    potential_final_base = matches[0]
                    target_mp4 = os.path.join(path, f"{potential_final_base}.mp4")
                    target_mp3 = os.path.join(path, f"{potential_final_base}.mp3")
                    
                    if os.path.exists(target_mp4) or os.path.exists(target_mp3):
                        log(f"Skipping repair for {temp_file}: corresponding final file exists.")
                        continue
                
                if os.path.getsize(temp_file) > 0:
                    os.remove(temp_file)
                    repaired_count += 1
                    log(f"Cleaned up incomplete file: {temp_file}")
            except Exception as e:
                log(f"Error repairing file {temp_file}: {str(e)}")

    return repaired_count

def resumeInterruptedDownloads():
    if not getINI("ResumeOnRestart"):
        return
    if not os.path.exists(StateFilePath):
        saveState([])
    with _global_state_lock:
        queue = loadState()
    downloads_to_resume = [item for item in queue if item.get("status") in ["running", "queued"]]
    if not downloads_to_resume:
        return
    path = getINI("ResultFolder") or DownloadPath
    if os.path.isdir(path):
        repaired = repairIncompleteFiles(path)
        log(f"Auto-repaired {repaired} files before resuming downloads")

    ui.message(_("Checking interrupted downloads..."))
    for item in downloads_to_resume:
        if YouTubeEXE in item["cmd"][0] and "--continue" not in item["cmd"]:
            item["cmd"].insert(1, "--continue")
        updateDownloadStatusInQueue(item.get("id"), "queued")
        if item.get("format") == "mp3":
            _cleanup_temp_files(item.get("path", ""), item.get("title", ""), item.get("format", ""))
    if not promptResumeDownloads(downloads_to_resume):
        for item in downloads_to_resume:
            updateDownloadStatusInQueue(item.get("id"), "cancelled")
        clearState()
        return
    for item in downloads_to_resume:
        updateDownloadStatusInQueue(item.get("id"), "queued")
        if item.get("format") == "mp3":
            _cleanup_temp_files(item.get("path", ""), item.get("title", ""), item.get("format", ""))
        _download_queue.put(item)

def start_worker_threads():
    global _num_workers
    _num_workers = getINI("MaxConcurrentDownloads")
    for _ in range(_num_workers):
        t = threading.Thread(target=worker_loop, daemon=True)
        t.start()

def shutdown_workers():
    for _ in range(_num_workers):
        _download_queue.put(None)

def worker_loop():
    while True:
        item = _download_queue.get()
        if item is None:
            break
        run_download(item)
        _download_queue.task_done()

def _process_next_download():
    # Placeholder for processing the next download in the queue
    # This will be called by uTubeTrim to ensure queue processing
    pass

def run_download(item):
    download_id = item["id"]
    cmd = item["cmd"]
    save_path = item["path"]
    url = item["url"]
    title = item["title"]
    file_format = item["format"]
    is_playlist = item.get("is_playlist", False)
    is_trimming = item.get("trimming", False)
    
    with _global_active_lock:
        global _global_active_downloads
        _global_active_downloads += 1
        if _global_active_downloads == 1:
            wx.CallAfter(startHeartbeat)
    
    updateDownloadStatusInQueue(download_id, "running")
    log(f"Starting download for ID: {download_id}")
    log(f"Command: {cmd}")
    
    process = None
    try:
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE
        
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=save_path,
            startupinfo=si,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        log(f"Process started with PID: {process.pid}")

        timeout = 1800
        try:
            stdout, stderr = process.communicate(timeout=timeout)
            return_code = process.returncode

            stdout_str = stdout.decode('utf-8', errors='ignore')
            stderr_str = stderr.decode('utf-8', errors='ignore')
            
            if return_code == 0:
                log(f"Download for ID {download_id} completed successfully.")
                log(f"STDOUT: {stdout_str}")
                PlayWave('complete')
                if getINI("SayDownloadComplete"):
                    wx.CallAfter(ui.message, _("Download complete"))
                updateDownloadStatusInQueue(download_id, "completed")
            else:
                log(f"Download for ID {download_id} failed with return code {return_code}.")
                log(f"STDOUT: {stdout_str}")
                log(f"STDERR: {stderr_str}")
                PlayWave('failed')
                wx.CallAfter(ui.message, _("Download failed"))
                updateDownloadStatusInQueue(download_id, "failed")
        except subprocess.TimeoutExpired:
            log(f"Download for ID {download_id} timed out after {timeout} seconds.")
            if process:
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
            PlayWave('failed')
            wx.CallAfter(ui.message, _("Download failed due to timeout"))
            updateDownloadStatusInQueue(download_id, "failed")
    except Exception as e:
        log(f"Error during download execution for ID {download_id}: {e}")
        if process:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
        PlayWave('failed')
        wx.CallAfter(ui.message, _("Download failed due to an error"))
        updateDownloadStatusInQueue(download_id, "failed")
    finally:
        if not is_trimming:
            _cleanup_temp_files(save_path, title, file_format)
        removeCompletedOrFailedDownloadsFromQueue()
        with _global_active_lock:
            _global_active_downloads -= 1
            if _global_active_downloads == 0:
                wx.CallAfter(stopHeartbeat)
        log(f"Download for ID {download_id} finished.")

def convertToMP(mpFormat, savePath, isPlaylist=False, url=None, title=None):
    if not isBrowser():
        ui.message(_("Browser required"))
        return
    if not createFolder(savePath):
        ui.message(_("Cannot create folder"))
        return
    if os.path.isdir(savePath):
        repaired = repairIncompleteFiles(savePath)
        log(f"Auto-repaired {repaired} files before new download")

    url = url or getCurrentDocumentURL()
    if not url:
        ui.message(_("URL not found"))
        return
    is_youtube_url = any(y in url for y in [".youtube.", "youtu.be", "youtube.com"])
    if is_youtube_url:
        video_title = getWebSiteTitle()
        sanitized_title = validFilename(video_title)
        if checkFileExists(savePath, sanitized_title, mpFormat):
            if mpFormat == "mp3" and os.path.exists(os.path.join(savePath, f"{sanitized_title}.mp4")):
                log(f"MP4 exists for {sanitized_title}, allowing MP3 download")
            else:
                ui.message(_("File exists"))
                return
        if not isPlaylist:
            parsed = urllib.parse.urlparse(url)
            query_params = urllib.parse.parse_qs(parsed.query)
            if 'list' in query_params:
                del query_params['list']
            if 'index' in query_params:
                del query_params['index']
            new_query = urllib.parse.urlencode(query_params, doseq=True)
            url = urllib.parse.urlunparse((
                parsed.scheme, parsed.netloc, parsed.path,
                parsed.params, new_query, parsed.fragment
            ))
            log("Removed playlist parameters from URL for single video download")

        if not os.path.exists(YouTubeEXE):
            ui.message(_("yt-dlp.exe missing"))
            return
        PlayWave("start")
        output = os.path.join(savePath, "%(title)s.%(ext)s")
        
        use_multipart = getINI("UseMultiPart") and os.path.exists(Aria2cEXE)
        connections = getINI("MultiPartConnections")
        
        if mpFormat == "mp3":
            cmd = [
                YouTubeEXE, "--no-playlist" if not isPlaylist else "--yes-playlist",
                "-x", "--audio-format", "mp3",
                "--audio-quality", str(getINI("MP3Quality")),
                "--ffmpeg-location", ConverterEXE,
                "-o", output, "--ignore-errors", "--no-warnings", url
            ]
            if use_multipart:
                aria2_args = f"-x{connections} -j{connections} -s{connections} -k1M --file-allocation=none --allow-overwrite=true --max-tries=0 --retry-wait=1"
                cmd.extend(["--external-downloader", Aria2cEXE, 
                            "--external-downloader-args", f"aria2c:{aria2_args}"])
        else:
            cmd = [
                YouTubeEXE, "--no-playlist" if not isPlaylist else "--yes-playlist",
                "-f", "bv*[ext=mp4]+ba[ext=m4a]/b[ext=mp4]/bv*+ba/b",
                "--remux-video", "mp4",
                "--ffmpeg-location", ConverterEXE,
                "-o", output, "--ignore-errors", "--no-warnings", url
            ]
            if use_multipart:
                aria2_args = f"-x{connections} -j{connections} -s{connections} -k1M --file-allocation=none --allow-overwrite=true --max-tries=0 --retry-wait=1"
                cmd.extend(["--external-downloader", Aria2cEXE, 
                            "--external-downloader-args", f"aria2c:{aria2_args}"])
            else:
                cmd.extend(["--concurrent-fragments", str(connections)])
        download_obj = {
            "url": url, "title": sanitized_title, "format": mpFormat,
            "path": savePath, "cmd": cmd, "is_playlist": isPlaylist
        }
        download_id = addDownloadToQueue(download_obj)
        _download_queue.put(download_obj)
    else:
        ext = getMultimediaURLExtension()
        ext = ext.lstrip(".")
        if ext and isValidMultimediaExtension(ext):
            if not os.path.exists(ConverterEXE):
                ui.message(_("Error: ffmpeg.exe not found."))
                return
            multimediaLinkURL = getLinkURL()
            linkName = getLinkName()
            if checkFileExists(savePath, linkName, mpFormat):
                ui.message(_("File already exists. Skipping download."))
                return
            if not multimediaLinkURL:
                ui.message(_("No valid multimedia link found."))
                return
            multimediaLinkName = os.path.join(savePath, validFilename(linkName) + "." + mpFormat)
            if mpFormat == "mp3":
                cmd = [
                    ConverterEXE, "-i", multimediaLinkURL,
                    "-c:a", "libmp3lame", "-b:a", f"{getINI('MP3Quality')}k",
                    "-map", "0:a", "-y", multimediaLinkName
                ]
            else:
                cmd = [
                    ConverterEXE, "-i", multimediaLinkURL,
                    "-c:v", "libx265", "-preset", "fast", "-crf", "23",
                    "-c:a", "copy", "-map", "0:v?", "-map", "0:a?",
                    "-y", multimediaLinkName
                ]
            ui.message(_("Adding link as {format} to download queue").format(format=mpFormat.upper()))
            PlayWave("start")
            download_obj = {
                "url": multimediaLinkURL, "title": linkName, "format": mpFormat,
                "path": savePath, "cmd": cmd, "is_playlist": False
            }
            download_id = addDownloadToQueue(download_obj)
            _download_queue.put(download_obj)
        else:
            ui.message(_("Not a YouTube video or valid multimedia link"))

def setSpeed(sp):
    speech.setSpeechOption("rate", sp)
    speech.speak(" ")
