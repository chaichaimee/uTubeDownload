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

AddOnSummary = "uTubeDownload"
AddOnName = "uTubeDownload"
AddOnPath = os.path.dirname(__file__)
ToolsPath = os.path.join(AddOnPath, "Tools")
SoundPath = os.path.join(AddOnPath, "sounds")
AppData = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming')
DownloadPath = None
sectionName = AddOnName
processID = None
_download_queue_lock = threading.Lock()
_current_download_thread = None
_heartbeat_thread = None
_heartbeat_active = False
Aria2cEXE = os.path.join(ToolsPath, "aria2c.exe")

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
YouTubeEXE = os.path.join(ToolsPath, "yt-dlp.exe")
ConverterEXE = os.path.join(ToolsPath, "ffmpeg.exe")
ConverterPath = ToolsPath

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
        return []
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
    queue = loadState()
    download_obj["id"] = str(uuid.uuid4())
    download_obj["start_time"] = datetime.datetime.now().isoformat()
    download_obj["status"] = "queued"
    queue.append(download_obj)
    saveState(queue)
    log(f"Added download to queue: ID {download_obj['id']}")
    return download_obj["id"]

def updateDownloadStatusInQueue(download_id, status):
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
    queue = loadState()
    new_queue = [item for item in queue if item.get("status") not in ["completed", "failed", "cancelled"]]
    if len(new_queue) < len(queue):
        saveState(new_queue)
        log(f"Removed {len(queue) - len(new_queue)} completed/failed downloads from queue")

def makePrintable(s):
    return "".join(c if c.isprintable() else " " for c in str(s))

def validFilename(s):
    return "".join(c if c not in ["/", "\\", ":", "*", "<", ">", "?", "\"", "|", "\n", "\r", "\t"] else "_" for c in s)

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
        except Exception:
            ui.message(_("Cannot create folder"))
            log(f"Failed to create folder: {folder}")
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

def checkFileExists(savePath, title, extension):
    if not getINI("SkipExisting"):
        return False

    sanitized_title = validFilename(title)
    filename = os.path.join(savePath, f"{sanitized_title}.{extension}")
    if os.path.exists(filename):
        log(f"File already exists: {filename}")
        return True

    temp_patterns = [
        f"{filename}.part", f"{filename}.ytdl", f"{filename}.temp", f"{filename}.download",
        f"{os.path.join(savePath, sanitized_title)}*.part",
        f"{os.path.join(savePath, sanitized_title)}*.ytdl",
        f"{os.path.join(savePath, sanitized_title)}*.temp",
        f"{os.path.join(savePath, sanitized_title)}*.download",
        f"{os.path.join(savePath, sanitized_title)}*.f*.tmp",
        f"{os.path.join(savePath, sanitized_title)}*.f*.mp4",
        f"{os.path.join(savePath, sanitized_title)}*.f*.webm",
        f"{os.path.join(savePath, sanitized_title)}*.f*.m4a"
    ]
    for pattern in temp_patterns:
        if glob.glob(pattern):
            log(f"Found temp file matching pattern {pattern}, not skipping.")
            return False
    return False

def safeMessageBox(message, title, style):
    return gui.messageBox(message, title, style)

def promptResumeDownloads(downloads_list):
    count = len(downloads_list)
    msg = _("Found {count} interrupted downloads\nResume all?").format(count=count)
    return safeMessageBox(msg, _("Resume downloads"), wx.YES_NO) == wx.YES

def _kill_ffmpeg_processes():
    try:
        for proc in psutil.process_iter(['name']):
            if proc.name().lower() in ['ffmpeg.exe', 'yt-dlp.exe', 'aria2c.exe']:
                log(f"Terminating process {proc.name()} with PID {proc.pid}")
                proc.terminate()
                try:
                    proc.wait(timeout=3)
                except psutil.TimeoutExpired:
                    log(f"Force killing process {proc.name()} with PID {proc.pid}")
                    proc.kill()
    except Exception as e:
        log(f"Error terminating processes: {e}")

def _cleanup_webm_files(save_path, title, file_format, check_count=2):
    if not title or not save_path:
        log(f"Webm cleanup skipped: title or path missing (title: {title}, path: {save_path})")
        return

    sanitized_title = validFilename(title)
    base_filename = os.path.join(save_path, sanitized_title)

    for _ in range(check_count):
        temp_patterns = [
            f"{base_filename}.webm",
            f"{base_filename}*.webm",
            f"{base_filename}*.f*.webm",
            os.path.join(save_path, f"*{sanitized_title}*.webm"),
            os.path.join(save_path, f"*{sanitized_title}*.f*.webm"),
            os.path.join(save_path, f"*{sanitized_title}*-[0-9]*.webm"),
            os.path.join(save_path, f"*{sanitized_title}*[a-zA-Z0-9_-]*.webm"),
            os.path.join(save_path, "*.webm")
        ]
        # Add .mp4 cleanup for MP3 mode only
        if file_format == "mp3":
            temp_patterns.extend([
                f"{base_filename}.mp4",
                f"{base_filename}*.mp4",
                f"{base_filename}*.f*.mp4",
                os.path.join(save_path, f"*{sanitized_title}*.mp4"),
                os.path.join(save_path, f"*{sanitized_title}*.f*.mp4"),
                os.path.join(save_path, f"*{sanitized_title}*-[0-9]*.mp4"),
                os.path.join(save_path, f"*{sanitized_title}*[a-zA-Z0-9_-]*.mp4"),
                os.path.join(save_path, "*.mp4")
            ])
        for pattern in temp_patterns:
            for temp_file in glob.glob(pattern):
                try:
                    if os.path.exists(temp_file) and (file_format != "mp4" or temp_file.lower() != f"{base_filename}.mp4".lower()):
                        os.remove(temp_file)
                        log(f"Removed temp file in cleanup_webm_files: {temp_file}")
                except Exception as e:
                    log(f"Error removing temp file {temp_file}: {e}")

def _process_next_download():
    global _current_download_thread

    with _download_queue_lock:
        removeCompletedOrFailedDownloadsFromQueue()
        queue = loadState()

        next_download = None
        for item in queue:
            if item.get("status") == "queued":
                next_download = item
                break

        if next_download and _current_download_thread is None:
            updateDownloadStatusInQueue(next_download.get("id"), "running")

            title = next_download.get("title", _("Unknown title"))
            file_format = next_download.get("format", _("unknown format"))
            url = next_download.get("url", _("unknown URL"))
            cmd = next_download.get("cmd")
            save_path = next_download.get("path")
            is_playlist = next_download.get("is_playlist", False)
            download_id = next_download.get("id")
            trimming = next_download.get("trimming", False)

            if not cmd or not save_path or not url or not title or not file_format or download_id is None:
                updateDownloadStatusInQueue(download_id, "failed")
                wx.CallAfter(_process_next_download)
                return

            _cleanup_webm_files(save_path, title, file_format)
            startHeartbeat()
            _current_download_thread = converterThread(cmd, save_path, url, title, file_format,
                                                       resume=False, is_playlist=is_playlist, download_id=download_id, trimming=trimming)
            _current_download_thread.daemon = True
            _current_download_thread.start()
            WaitThread(_current_download_thread).start()

        elif not next_download:
            removeCompletedOrFailedDownloadsFromQueue()

def _on_download_complete(download_id, status):
    global _current_download_thread

    with _download_queue_lock:
        updateDownloadStatusInQueue(download_id, status)
        _current_download_thread = None
        stopHeartbeat()
        if status == "completed":
            PlayWave("complete")
            queue = loadState()
            for item in queue:
                if item.get("id") == download_id:
                    if item.get("trimming", False):
                        wx.CallAfter(ui.message, _("Trim complete"))
                    else:
                        _cleanup_temp_files_immediately(
                            item.get("title", ""),
                            item.get("path", ""),
                            item.get("format", "")
                        )
                    break
        elif status == "failed":
            queue = loadState()
            for item in queue:
                if item.get("id") == download_id:
                    if item.get("trimming", False):
                        wx.CallAfter(ui.message, _("Trim failed"))
                    else:
                        wx.CallAfter(ui.message, _("Download failed"))
                    break

    winsound.PlaySound(None, winsound.SND_PURGE)
    wx.CallAfter(_process_next_download)

def _cleanup_temp_files_immediately(title, path, file_format):
    if not title or not path:
        log(f"Temp cleanup skipped: title or path missing (title: {title}, path: {path})")
        return

    sanitized_title = validFilename(title)
    base_filename = os.path.join(path, sanitized_title)

    temp_patterns = [
        f"{base_filename}.part", f"{base_filename}.ytdl", f"{base_filename}.temp", f"{base_filename}.download",
        f"{base_filename}*.f*.tmp", f"{base_filename}*.f*.webm", f"{base_filename}*.f*.m4a",
        f"{base_filename}.webm", f"{base_filename}.m4a", f"{base_filename}.opus",
        f"{base_filename}.aria2", f"{base_filename}.part.aria2",
        f"{base_filename}*.aria2", f"{base_filename}*.part.aria2",
        f"{base_filename}*-[0-9]*.webm", f"{base_filename}*[a-zA-Z0-9_-]*.webm", os.path.join(path, "*.webm"),
        f"{base_filename}*.f*.m4a.part.aria2", f"{base_filename}*.f*.webm.part.aria2"
    ]
    # Add .mp4 cleanup for MP3 mode only
    if file_format == "mp3":
        temp_patterns.extend([
            f"{base_filename}.mp4",
            f"{base_filename}*.mp4",
            f"{base_filename}*.f*.mp4",
            os.path.join(path, f"*{sanitized_title}*.mp4"),
            os.path.join(path, f"*{sanitized_title}*.f*.mp4"),
            os.path.join(path, f"*{sanitized_title}*-[0-9]*.mp4"),
            os.path.join(path, f"*{sanitized_title}*[a-zA-Z0-9_-]*.mp4"),
            os.path.join(path, "*.mp4"),
            f"{base_filename}*.f*.mp4.part.aria2",
            os.path.join(path, "*.mp4.part.aria2")
        ])
    if "is_playlist" in loadState() and loadState() and loadState()[0].get("is_playlist", False):
        playlist_temp_patterns = [
            os.path.join(path, f"*{sanitized_title}*.part"), os.path.join(path, f"*{sanitized_title}*.ytdl"),
            os.path.join(path, f"*{sanitized_title}*.temp"), os.path.join(path, f"*{sanitized_title}*.download"),
            os.path.join(path, f"*{sanitized_title}*.f*.tmp"), os.path.join(path, f"*{sanitized_title}*.f*.webm"),
            os.path.join(path, f"*{sanitized_title}*.f*.m4a"),
            os.path.join(path, f"*{sanitized_title}.webm"), os.path.join(path, f"*{sanitized_title}.m4a"),
            os.path.join(path, f"*{sanitized_title}.opus"), os.path.join(path, f"*{sanitized_title}*-[0-9]*.webm"),
            os.path.join(path, f"*{sanitized_title}*[a-zA-Z0-9_-]*.webm"), os.path.join(path, "*.webm"),
            os.path.join(path, f"*{sanitized_title}*.part.aria2"), os.path.join(path, f"*{sanitized_title}*.aria2"),
            os.path.join(path, f"*{sanitized_title}*.f*.m4a.part.aria2"),
            os.path.join(path, f"*{sanitized_title}*.f*.webm.part.aria2")
        ]
        if file_format == "mp3":
            playlist_temp_patterns.extend([
                os.path.join(path, f"*{sanitized_title}*.mp4"),
                os.path.join(path, f"*{sanitized_title}*.f*.mp4"),
                os.path.join(path, f"*{sanitized_title}*-[0-9]*.mp4"),
                os.path.join(path, f"*{sanitized_title}*[a-zA-Z0-9_-]*.mp4"),
                os.path.join(path, f"*{sanitized_title}*.f*.mp4.part.aria2")
            ])
        temp_patterns.extend(playlist_temp_patterns)
    generic_temp_patterns = [
        os.path.join(path, "*.part"), os.path.join(path, "*.ytdl"), os.path.join(path, "*.temp"), os.path.join(path, "*.download"),
        os.path.join(path, "*.f*.tmp"), os.path.join(path, "*.f*.webm"), os.path.join(path, "*.f*.m4a"),
        os.path.join(path, "*.webm"), os.path.join(path, "*.part.aria2"), os.path.join(path, "*.aria2"),
        os.path.join(path, "*.f*.m4a.part.aria2"), os.path.join(path, "*.f*.webm.part.aria2")
    ]
    if file_format == "mp3":
        generic_temp_patterns.extend([
            os.path.join(path, "*.mp4"),
            os.path.join(path, "*.f*.mp4"),
            os.path.join(path, "*.mp4.part.aria2")
        ])
    temp_patterns.extend(generic_temp_patterns)
    unique_temp_patterns = list(set(temp_patterns))
    for pattern in unique_temp_patterns:
        for temp_file in glob.glob(pattern):
            try:
                final_target_file_pattern = os.path.join(path, f"{sanitized_title}.{file_format}")
                if file_format != "mp4" or temp_file.lower() != final_target_file_pattern.lower():
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        log(f"Removed temp file: {temp_file}")
            except Exception as e:
                log(f"Error removing temp file {temp_file}: {e}")

def repairIncompleteFiles(path):
    repaired_count = 0
    patterns = [
        "*.part", "*.ytdl", "*.temp", "*.download", "*.f*.tmp", "*.f*.mp4", "*.f*.webm", "*.f*.m4a",
        "*.part.aria2", "*.aria2", "*.f*.m4a.part.aria2", "*.f*.mp4.part.aria2", "*.f*.webm.part.aria2"
    ]
    
    for pattern in patterns:
        for temp_file in glob.glob(os.path.join(path, pattern)):
            try:
                base_name = os.path.splitext(temp_file)[0]
                if os.path.exists(base_name + ".mp4"):
                    target = base_name + ".mp4"
                elif os.path.exists(base_name + ".mp3"):
                    target = base_name + ".mp3"
                else:
                    continue
                
                if os.path.getsize(temp_file) > 0:
                    shutil.copy(temp_file, target)
                    os.remove(temp_file)
                    repaired_count += 1
                    log(f"Repaired incomplete file: {temp_file} -> {target}")
            except Exception as e:
                log(f"Error repairing file {temp_file}: {str(e)}")
    
    return repaired_count

def resumeInterruptedDownloads():
    if not getINI("ResumeOnRestart"):
        return
    if not os.path.exists(StateFilePath):
        saveState([])
    queue = loadState()
    downloads_to_resume = [item for item in queue if item.get("status") in ["running", "queued"]]
    if not downloads_to_resume:
        return
    
    # Auto-repair before resuming
    path = getINI("ResultFolder") or DownloadPath
    if os.path.isdir(path):
        repaired = repairIncompleteFiles(path)
        log(f"Auto-repaired {repaired} files before resuming downloads")
    
    ui.message(_("Checking interrupted downloads..."))
    _kill_ffmpeg_processes()
    for item in downloads_to_resume:
        if YouTubeEXE in item["cmd"][0] and "--continue" not in item["cmd"]:
            item["cmd"].insert(1, "--continue")
        updateDownloadStatusInQueue(item.get("id"), "queued")
        if item.get("format") == "mp3":
            _cleanup_webm_files(item.get("path", ""), item.get("title", ""), item.get("format", ""))
    if not promptResumeDownloads(downloads_to_resume):
        for item in downloads_to_resume:
            updateDownloadStatusInQueue(item.get("id"), "cancelled")
        clearState()
        return
    for item in downloads_to_resume:
        if YouTubeEXE in item["cmd"][0] and "--continue" not in item["cmd"]:
            item["cmd"].insert(1, "--continue")
        updateDownloadStatusInQueue(item.get("id"), "queued")
        if item.get("format") == "mp3":
            _cleanup_webm_files(item.get("path", ""), item.get("title", ""), item.get("format", ""))
    wx.CallAfter(_process_next_download)

def convertToMP(mpFormat, savePath, isPlaylist=False, url=None, title=None):
    if not isBrowser():
        ui.message(_("Browser required"))
        return
    if not createFolder(savePath):
        ui.message(_("Cannot create folder"))
        return
    
    # Auto-repair before new download
    if os.path.isdir(savePath):
        repaired = repairIncompleteFiles(savePath)
        log(f"Auto-repaired {repaired} files before new download")
    
    url = url or getCurrentDocumentURL()
    if not url:
        ui.message(_("URL not found"))
        return
    is_youtube_url = any(y in url for y in [".youtube.", "youtu.be", "youtube.com"])
    if is_youtube_url:
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
        title = title or getWebSiteTitle()
        if checkFileExists(savePath, title, mpFormat):
            ui.message(_("File exists"))
            return
        ui.message(_("Download {format}").format(format=mpFormat.upper()))
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
                "-o", output, url
            ]
            if use_multipart:
                aria2_args = f"-x{connections} -j{connections} -s{connections} -k1M --file-allocation=none --allow-overwrite=true --max-tries=0 --retry-wait=1"
                cmd.extend(["--external-downloader", Aria2cEXE, 
                            "--external-downloader-args", aria2_args])
        else:
            cmd = [
                YouTubeEXE, "--no-playlist" if not isPlaylist else "--yes-playlist",
                "-f", "bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best",
                "--merge-output-format", "mp4",
                "--ffmpeg-location", ConverterEXE,
                "-o", output, url
            ]
            if use_multipart:
                aria2_args = f"-x{connections} -j{connections} -s{connections} -k1M --file-allocation=none --allow-overwrite=true --max-tries=0 --retry-wait=1"
                cmd.extend(["--external-downloader", Aria2cEXE, 
                            "--external-downloader-args", aria2_args])
            else:
                cmd.extend(["--concurrent-fragments", str(connections)])
        download_id = addDownloadToQueue({
            "url": url, "title": title, "format": mpFormat,
            "path": savePath, "cmd": cmd, "is_playlist": isPlaylist
        })
        wx.CallAfter(_process_next_download)
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
            initial_state = {
                "url": multimediaLinkURL, "title": linkName, "format": mpFormat,
                "path": savePath, "cmd": cmd, "is_playlist": False
            }
            download_id = addDownloadToQueue(initial_state)
            wx.CallAfter(_process_next_download)
        else:
            ui.message(_("Not a YouTube video or valid multimedia link"))

class converterThread(threading.Thread):
    def __init__(self, cmd, path, url, title, file_format, download_id=None, resume=False, is_playlist=False, trimming=False):
        super().__init__()
        self.cmd = cmd
        self.path = path
        self.url = url
        self.title = title
        self.file_format = file_format
        self.download_id = download_id
        self.resume = resume
        self.is_playlist = is_playlist
        self.trimming = trimming
        self.daemon = True
        self.process = None

    def run(self):
        global processID
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        try:
            if not os.path.isdir(self.path):
                _on_download_complete(self.download_id, "failed")
                return
            
            if self.trimming:
                wx.CallAfter(tones.beep, 1000, 100)
            else:
                PlayWave("start")
            
            self.process = subprocess.Popen(
                self.cmd,
                cwd=self.path,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                startupinfo=si,
                creationflags=subprocess.CREATE_NO_WINDOW,
                encoding="utf-8",
                text=True,
                errors='ignore'
            )
            processID = self.process.pid
            
            stdout, stderr = self.process.communicate()
            
            if stdout:
                log(f"Process stdout: {stdout[:500]}")
            if stderr:
                log(f"Process stderr: {stderr[:500]}")
            
            if self.file_format == "mp3":
                _kill_ffmpeg_processes()
                _cleanup_webm_files(self.path, self.title, self.file_format)
            
            if self.process.returncode == 0:
                if self.trimming:
                    wx.CallAfter(tones.beep, 2000, 100)
                _on_download_complete(self.download_id, "completed")
            else:
                if self.trimming:
                    wx.CallAfter(tones.beep, 500, 300)
                _on_download_complete(self.download_id, "failed")
        except Exception as e:
            log(f"Download exception: {str(e)}")
            if self.trimming:
                wx.CallAfter(tones.beep, 500, 300)
            _on_download_complete(self.download_id, "failed")
        finally:
            processID = None

class WaitThread(threading.Thread):
    def __init__(self, targetThread):
        super().__init__()
        self.target = targetThread
        self.daemon = True
    def run(self):
        while self.target.is_alive():
            time.sleep(0.5)
        self.target.join()
