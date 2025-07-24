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
import psutil
from scriptHandler import script
import gui

AddOnSummary = "uTubeDownload"
AddOnName = "uTubeDownload"
AddOnPath = os.path.dirname(__file__)
ToolsPath = os.path.join(AddOnPath, "Tools")
SoundPath = os.path.join(AddOnPath, "sounds")
AppData = os.environ["APPDATA"]
DownloadPath = os.environ["APPDATA"] # Default to AppData if not configured
sectionName = AddOnName
processID = None
_download_queue_lock = threading.Lock()
_current_download_thread = None
_heartbeat_thread = None
_heartbeat_active = False

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
        time.sleep(4)  # Wait for the 4-second sound to finish before playing again

def startHeartbeat():
    global _heartbeat_thread, _heartbeat_active
    if not _heartbeat_active:
        _heartbeat_active = True
        _heartbeat_thread = threading.Thread(target=_heartbeat_loop, daemon=True)
        _heartbeat_thread.start()

def stopHeartbeat():
    global _heartbeat_thread, _heartbeat_active
    _heartbeat_active = False
    if _heartbeat_thread:
        _heartbeat_thread.join()
    winsound.PlaySound(None, winsound.SND_PURGE)

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
    download_obj["id"] = int(time.time() * 1000)
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
        # Always log to NVDA for debugging
        api.log.info(f"uTubeDownload: {makePrintable(s)}")
        api.log.debug(f"uTubeDownload Debug: {makePrintable(s)}")
        # Log to file only if Logging is enabled
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
    if not obj.treeInterceptor:
        ui.message(_("No browser found"))
        return False
    return True

def getCurrentDocumentURL():
    try:
        obj = api.getFocusObject()
        if hasattr(obj, 'treeInterceptor') and obj.treeInterceptor:
            url = obj.treeInterceptor.documentConstantIdentifier
            if url and ("/shorts/" in url or "youtube.com/shorts/" in url):
                video_id = url.split("/shorts/")[-1].split("?")[0]
                url = f"https://www.youtube.com/watch?v={video_id}"
            return urllib.parse.unquote(url) if url else None
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
        for suffix in [" - YouTube", "| YouTube"]:
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
            if proc.name().lower() in ['ffmpeg.exe', 'yt-dlp.exe']:
                log(f"Terminating process {proc.name()} with PID {proc.pid}")
                proc.terminate()
                try:
                    proc.wait(timeout=3)  # Wait up to 3 seconds for termination
                except psutil.TimeoutExpired:
                    log(f"Force killing process {proc.name()} with PID {proc.pid}")
                    proc.kill()
    except Exception as e:
        log(f"Error terminating processes: {e}")

def _cleanup_webm_files(savePath, title, file_format, check_count=2):
    if not title or not savePath:
        log(f"Webm cleanup skipped: title or path missing (title: {title}, path: {savePath})")
        return
    
    sanitized_title = validFilename(title)
    base_filename = os.path.join(savePath, sanitized_title)
    
    for _ in range(check_count):
        webm_patterns = [
            f"{base_filename}.webm",
            f"{base_filename}*.webm",
            f"{base_filename}*.f*.webm",
            os.path.join(savePath, f"*{sanitized_title}*.webm"),
            os.path.join(savePath, f"*{sanitized_title}*.f*.webm"),
            os.path.join(savePath, f"*{sanitized_title}*-[0-9]*.webm"),
            os.path.join(savePath, f"*{sanitized_title}*[a-zA-Z0-9_-]*.webm"),
            os.path.join(savePath, "*.webm")  # Aggressive cleanup
        ]
        
        for pattern in webm_patterns:
            for webm_file in glob.glob(pattern):
                try:
                    if os.path.exists(webm_file) and webm_file.lower() != f"{base_filename}.{file_format}".lower():
                        log(f"Attempting to delete webm file: {webm_file}")
                        os.remove(webm_file)
                        log(f"Successfully deleted webm file: {webm_file}")
                except Exception as e:
                    log(f"Error deleting webm file {webm_file}: {e}")

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
            format = next_download.get("format", _("unknown format"))
            url = next_download.get("url", _("unknown URL"))
            cmd = next_download.get("cmd")
            save_path = next_download.get("path")
            is_playlist = next_download.get("is_playlist", False)
            download_id = next_download.get("id")

            if not cmd or not save_path or not url or not title or not format or download_id is None:
                updateDownloadStatusInQueue(download_id, "failed")
                log(f"Skipping invalid download entry: {next_download}")
                wx.CallAfter(_process_next_download)
                return

            _cleanup_webm_files(save_path, title, format)
            log(f"Starting download: {title} (ID: {download_id})")
            startHeartbeat()

            _current_download_thread = converterThread(cmd, save_path, url, title, format,
                                                     resume=False, is_playlist=is_playlist, download_id=download_id)
            _current_download_thread.daemon = True
            _current_download_thread.start()
            WaitThread(_current_download_thread).start()
        elif not next_download:
            log("No more downloads in queue.")
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
                    _cleanup_temp_files_immediately(
                        item.get("title", ""),
                        item.get("path", ""),
                        item.get("format", "")
                    )
                    _cleanup_webm_files(
                        item.get("path", ""),
                        item.get("title", ""),
                        item.get("format", "")
                    )
                    break
        wx.CallAfter(_process_next_download)

def _cleanup_temp_files_immediately(title, path, file_format):
    if not title or not path:
        log(f"Cleanup skipped: title or path missing (title: {title}, path: {path})")
        return
    
    sanitized_title = validFilename(title)
    base_filename = os.path.join(path, sanitized_title)

    log(f"Starting cleanup for '{base_filename}' (format: {file_format})")

    temp_patterns = [
        f"{base_filename}.part",
        f"{base_filename}.ytdl",
        f"{base_filename}.temp",
        f"{base_filename}.download",
        f"{base_filename}*.f*.tmp",
        f"{base_filename}*.f*.mp4",
        f"{base_filename}*.f*.webm",
        f"{base_filename}*.f*.m4a",
        f"{base_filename}.webm",
        f"{base_filename}.m4a",
        f"{base_filename}.opus",
        f"{base_filename}*-[0-9]*.webm",
        f"{base_filename}*[a-zA-Z0-9_-]*.webm",
        os.path.join(path, "*.webm")  # Aggressive cleanup
    ]

    if "is_playlist" in loadState()[0] and loadState()[0]["is_playlist"]:
        playlist_temp_patterns = [
            os.path.join(path, f"*{validFilename(title)}*.part"),
            os.path.join(path, f"*{validFilename(title)}*.ytdl"),
            os.path.join(path, f"*{validFilename(title)}*.temp"),
            os.path.join(path, f"*{validFilename(title)}*.download"),
            os.path.join(path, f"*{validFilename(title)}*.f*.tmp"),
            os.path.join(path, f"*{validFilename(title)}*.f*.mp4"),
            os.path.join(path, f"*{validFilename(title)}*.f*.webm"),
            os.path.join(path, f"*{validFilename(title)}*.f*.m4a"),
            os.path.join(path, f"*{validFilename(title)}.webm"),
            os.path.join(path, f"*{validFilename(title)}.m4a"),
            os.path.join(path, f"*{validFilename(title)}.opus"),
            os.path.join(path, f"*{validFilename(title)}*-[0-9]*.webm"),
            os.path.join(path, f"*{validFilename(title)}*[a-zA-Z0-9_-]*.webm"),
            os.path.join(path, "*.webm")
        ]
        temp_patterns.extend(playlist_temp_patterns)
    
    generic_temp_patterns = [
        os.path.join(path, "*.part"),
        os.path.join(path, "*.ytdl"),
        os.path.join(path, "*.temp"),
        os.path.join(path, "*.download"),
        os.path.join(path, "*.f*.tmp"),
        os.path.join(path, "*.f*.mp4"),
        os.path.join(path, "*.f*.webm"),
        os.path.join(path, "*.f*.m4a"),
        os.path.join(path, "*.webm")
    ]
    temp_patterns.extend(generic_temp_patterns)

    unique_temp_patterns = list(set(temp_patterns))

    for pattern in unique_temp_patterns:
        for temp_file in glob.glob(pattern):
            try:
                final_target_file_pattern = os.path.join(path, f"{sanitized_title}.{file_format}")
                if temp_file.lower() == final_target_file_pattern.lower():
                    log(f"Skipping deletion of final target file: {temp_file}")
                    continue

                if os.path.exists(temp_file):
                    log(f"Attempting to delete temporary file: {temp_file}")
                    os.remove(temp_file)
                    log(f"Successfully deleted temporary file: {temp_file}")
            except Exception as e:
                log(f"Error deleting temp file {temp_file}: {e}")

def resumeInterruptedDownloads():
    if not getINI("ResumeOnRestart"):
        log("ResumeOnRestart is disabled, skipping resume")
        return
    
    if not os.path.exists(StateFilePath):
        saveState([])
        log("State file not found, created empty state")
    
    queue = loadState()
    downloads_to_resume = [item for item in queue if item.get("status") in ["running", "queued"]]
    if not downloads_to_resume:
        log("No interrupted downloads found")
        return

    ui.message(_("Checking interrupted downloads..."))
    log(f"Found {len(downloads_to_resume)} interrupted downloads")
    # Terminate any lingering ffmpeg or yt-dlp processes
    _kill_ffmpeg_processes()
    # Clean up all .webm files in the download directory before resuming
    for item in downloads_to_resume:
        if item.get("format") == "mp3":
            _cleanup_webm_files(item.get("path", ""), item.get("title", ""), item.get("format", ""))
    
    if not promptResumeDownloads(downloads_to_resume):
        for item in downloads_to_resume:
            updateDownloadStatusInQueue(item.get("id"), "cancelled")
        clearState()
        log("User cancelled resuming downloads")
        return

    for item in downloads_to_resume:
        if YouTubeEXE in item["cmd"][0] and "--continue" not in item["cmd"]:
            item["cmd"].insert(1, "--continue")
        updateDownloadStatusInQueue(item.get("id"), "queued")
        log(f"Resuming download: {item.get('title', 'Unknown')} (ID: {item.get('id')})")
        # Clean up again after setting to queued
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
    url = url or getCurrentDocumentURL()
    if not url:
        ui.message(_("URL not found"))
        return

    is_youtube_url = any(y in url for y in [".youtube.", "youtu.be", "youtube.com", "https://www.youtube.com/watch?v="])

    if is_youtube_url:
        if not os.path.exists(YouTubeEXE):
            ui.message(_("yt-dlp.exe missing"))
            return

        title = title or getWebSiteTitle()
        if checkFileExists(savePath, title, mpFormat):
            ui.message(_("File exists"))
            return
            
        ui.message(_("Download {format}").format(format=mpFormat.upper()))
        PlayWave("start")
        log(f"Initiating download: {title} as {mpFormat}")

        output = os.path.join(savePath, "%(title)s.%(ext)s")
        if mpFormat == "mp3":
            cmd = [
                YouTubeEXE,
                "--no-playlist" if not isPlaylist else "--yes-playlist",
                "-x", "--audio-format", "mp3",
                "--audio-quality", str(getINI("MP3Quality")),
                "--ffmpeg-location", ConverterEXE,
                "-o", output, url
            ]
        else:
            cmd = [
                YouTubeEXE,
                "--no-playlist" if not isPlaylist else "--yes-playlist",
                "-f", "bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best",
                "--merge-output-format", "mp4",
                "--ffmpeg-location", ConverterEXE,
                "-o", output, url
            ]
            
        download_id = addDownloadToQueue({
            "url": url, 
            "title": title, 
            "format": mpFormat,
            "path": savePath, 
            "cmd": cmd, 
            "is_playlist": isPlaylist
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

            multimediaLinkName = os.path.join(savePath, linkName + "." + mpFormat)

            if mpFormat == "mp3":
                cmd = [
                    ConverterEXE,
                    "-i", multimediaLinkURL,
                    "-c:a", "libmp3lame",
                    "-b:a", f"{getINI('MP3Quality')}k",
                    "-map", "0:a",
                    "-y",
                    multimediaLinkName
                ]
            else:
                cmd = [
                    ConverterEXE,
                    "-i", multimediaLinkURL,
                    "-c:v", "copy",
                    "-c:a", "copy",
                    "-map", "0:v?",
                    "-map", "0:a?",
                    "-y",
                    multimediaLinkName
                ]
            cmd = [x for x in cmd if x is not None]
            ui.message(_("Adding link as {format} to download queue").format(format=mpFormat.upper()))
            PlayWave("start")
            log(f"Initiating multimedia download: {linkName} as {mpFormat}")

            initial_state = {
                "url": multimediaLinkURL,
                "title": linkName,
                "format": mpFormat,
                "path": savePath,
                "cmd": cmd,
                "is_playlist": False
            }
            download_id = addDownloadToQueue(initial_state)
            wx.CallAfter(_process_next_download)
        else:
            ui.message(_("Not a YouTube video or valid multimedia link"))
            log("Invalid URL for download")

class converterThread(threading.Thread):
    def __init__(self, cmd, path, url, title, format, download_id=None, resume=False, is_playlist=False):
        super().__init__()
        self.cmd = cmd
        self.path = path
        self.url = url
        self.title = title
        self.format = format
        self.download_id = download_id
        self.resume = resume
        self.is_playlist = is_playlist
        self.daemon = True
        self.process = None

    def run(self):
        global processID
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        try:
            if not os.path.isdir(self.path):
                _on_download_complete(self.download_id, "failed")
                ui.message(_("Error: Download folder missing. Cannot start download."))
                log(f"Download folder missing: {self.path}")
                return

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
            log(f"Started process with PID: {processID}")

            stdout, stderr = self.process.communicate()

            # Terminate any lingering ffmpeg or yt-dlp processes
            if self.format == "mp3":
                _kill_ffmpeg_processes()
                _cleanup_webm_files(self.path, self.title, self.format)

            if self.process.returncode == 0:
                expected_file_stem = os.path.join(self.path, validFilename(self.title))
                if self._verify_final_file(expected_file_stem):
                    _on_download_complete(self.download_id, "completed")
                    log(f"Download completed successfully for ID {self.download_id}")
                else:
                    _on_download_complete(self.download_id, "failed")
                    log(f"Final file verification failed for ID {self.download_id}. Output: {stdout}\nErrors: {stderr}")
            else:
                _on_download_complete(self.download_id, "failed")
                log(f"Download failed for ID {self.download_id}. Return code: {self.process.returncode}\nOutput: {stdout}\nErrors: {stderr}")
        except Exception as e:
            _on_download_complete(self.download_id, "failed")
            log(f"Download error for ID {self.download_id}: {str(e)}")
        finally:
            processID = None
            log(f"Process ended for ID {self.download_id}")

    def _verify_final_file(self, expected_filepath_stem):
        search_pattern = f"{expected_filepath_stem}*.{self.format}"
        found_files = glob.glob(search_pattern)

        if not found_files:
            log(f"No final file found matching pattern: {search_pattern}")
            return False

        for filepath in found_files:
            try:
                if os.path.exists(filepath) and os.path.getsize(filepath) >= 1024:
                    log(f"Verified final file: {filepath}")
                    return True
            except Exception:
                continue
        log(f"No valid final file found for pattern: {search_pattern}")
        return False

class WaitThread(threading.Thread):
    def __init__(self, targetThread):
        super().__init__()
        self.target = targetThread
        self.daemon = True

    def run(self):
        while self.target.is_alive():
            time.sleep(0.5)
        self.target.join()
