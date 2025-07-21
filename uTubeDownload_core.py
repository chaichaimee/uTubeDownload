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
from gui import guiHelper
import gui

# Constants
AddOnSummary = "uTubeDownload"
AddOnName = "uTubeDownload"
AddOnPath = os.path.dirname(__file__)
ToolsPath = os.path.join(AddOnPath, "Tools")
SoundPath = os.path.join(AddOnPath, "sounds")
AppData = os.environ["APPDATA"]
DownloadPath = os.path.join(AppData, "uTubeDownload")
sectionName = AddOnName
processID = None

def getStateFilePath():
    try:
        import globalVars
        if globalVars.appArgs.secure:
            return os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'nvda', 'uTubeDownload.json')
        
        configDir = globalVars.appArgs.configPath
        if not configDir:
            configDir = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'nvda')
        
        return os.path.join(configDir, 'uTubeDownload.json')
    except Exception:
        return os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'nvda', 'uTubeDownload.json')

StateFilePath = getStateFilePath()

# Download manager state
_current_download_thread = None
_download_queue_lock = threading.Lock()

MultimediaExtensions = {
    "aac", "avi", "flac", "mkv", "m3u8", "m4a", "m4s", "m4v",
    "mpg", "mov", "mp2", "mp3", "mp4", "mpeg", "mpegts", "ogg",
    "ogv", "oga", "ts", "vob", "wav", "webm", "wmv", "f4v",
    "flv", "swf", "avchd", "3gp"
}

invalidCharactersForFilename = ["/", "\\", ":", "*", "<", ">", "?", "!", "|"]

YouTubeEXE = os.path.join(ToolsPath, "yt-dlp.exe")
ConverterEXE = os.path.join(ToolsPath, "ffmpeg.exe")

def getINI(key):
    return config.conf[sectionName][key]

def setINI(key, value):
    config.conf[sectionName][key] = value

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
                state_data = json.load(f)
                return state_data if isinstance(state_data, list) else []
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

def removeCompletedOrFailedDownloadsFromQueue():
    queue = loadState()
    new_queue = [item for item in queue if item.get("status") not in ["completed", "failed", "cancelled"]]
    if len(new_queue) < len(queue):
        saveState(new_queue)

def makePrintable(s):
    return "".join(c if c.isprintable() else " " for c in str(s))

def validFilename(s):
    return "".join(c if c not in invalidCharactersForFilename else "_" for c in s)

def log(s):
    if getINI("Logging"):
        try:
            path = getINI("ResultFolder") or DownloadPath
            os.makedirs(path, exist_ok=True)
            with open(os.path.join(path, "log.txt"), "a", encoding="utf-8") as f:
                f.write(f"{datetime.datetime.now()} - {makePrintable(s)}\n")
        except Exception:
            pass

def PlayWave(filename):
    path = os.path.join(SoundPath, filename + ".wav")
    if os.path.exists(path) and getINI("BeepWhileConverting"):
        try:
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception:
            pass

def createFolder(folder):
    if not os.path.isdir(folder):
        try:
            os.makedirs(folder, exist_ok=True)
            return True
        except Exception:
            ui.message(_("Cannot create folder"))
            return False
    return True

def CheckFolders():
    global DownloadPath
    folder = getINI("ResultFolder") or os.path.join(AppData, "uTubeDownload")
    setINI("ResultFolder", folder)
    DownloadPath = folder
    createFolder(DownloadPath)
    createFolder(ToolsPath)
    createFolder(SoundPath)
    if not os.path.exists(YouTubeEXE):
        ui.message(_("yt-dlp.exe not found"))
    if not os.path.exists(ConverterEXE):
        ui.message(_("ffmpeg.exe not found"))

    if not os.path.exists(StateFilePath):
        saveState([])

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
    if obj.role == controlTypes.ROLE_LINK:
        url = obj.value
        if url:
            url = urllib.parse.unquote(url)
            return url[:-1] if url.endswith("/") else url
    return ""

def getLinkName():
    obj = api.getNavigatorObject()
    if obj.role == controlTypes.ROLE_LINK:
        return validFilename(obj.name)
    return ""

def getMultimediaURLExtension():
    url = getLinkURL()
    return url[url.rfind("."):].lower() if "." in url else ""

def isValidMultimediaExtension(ext):
    return ext.replace(".", "") in MultimediaExtensions

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
    filename = os.path.join(savePath, f"{title}.{extension}")
    if os.path.exists(filename):
        return True
    for pattern in [f"{filename}.part", f"{filename}.ytdl", f"{filename}.temp", f"{filename}.download"]:
        if os.path.exists(pattern):
            return False
    return False

def safeMessageBox(message, title, style):
    return gui.messageBox(message, title, style)

def promptResumeDownloads(downloads_list):
    count = len(downloads_list)
    msg = _("Found {count} interrupted downloads\nResume all?").format(count=count)
    return safeMessageBox(msg, _("Resume downloads"), wx.YES_NO) == wx.YES

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

            log(f"Starting sequential download: {title} (ID: {download_id})")

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
        wx.CallAfter(_process_next_download)

def resumeInterruptedDownloads():
    if not getINI("ResumeOnRestart"):
        return
    
    if not os.path.exists(StateFilePath):
        saveState([])
    
    queue = loadState()
    downloads_to_resume = [item for item in queue if item.get("status") in ["running", "queued"]]
    if not downloads_to_resume:
        return

    ui.message(_("Checking interrupted downloads..."))
    if not promptResumeDownloads(downloads_to_resume):
        for item in downloads_to_resume:
            updateDownloadStatusInQueue(item.get("id"), "cancelled")
        return

    for item in downloads_to_resume:
        if YouTubeEXE in item["cmd"][0] and "--continue" not in item["cmd"]:
            item["cmd"].insert(1, "--continue")
        updateDownloadStatusInQueue(item.get("id"), "queued")

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

    is_youtube_url = ".youtube." in url or "http://googleusercontent.com/youtube.com/" in url

    if is_youtube_url:
        if not os.path.exists(YouTubeEXE):
            ui.message(_("yt-dlp.exe missing"))
            return

        title = title or getWebSiteTitle()
        if checkFileExists(savePath, title, mpFormat):
            ui.message(_("File exists"))
            return
        ui.message(_("Adding {format} to download queue").format(format=mpFormat.upper()))

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
            multimediaLinkName = os.path.join(savePath, linkName + "." + mpFormat)

            if checkFileExists(savePath, linkName, mpFormat):
                ui.message(_("File already exists. Skipping download."))
                return

            if not multimediaLinkURL:
                ui.message(_("No valid multimedia link found."))
                return

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
                return

            self.process = subprocess.Popen(
                self.cmd,
                cwd=self.path,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                startupinfo=si,
                encoding="utf-8",
                text=True
            )
            processID = self.process.pid

            stdout, stderr = self.process.communicate()

            if self.process.returncode == 0:
                expected_file_stem = os.path.join(self.path, validFilename(self.title))
                if self._verify_final_file(expected_file_stem):
                    self._cleanup_temp_files(force=True)
                    _on_download_complete(self.download_id, "completed")
                    PlayWave("complete")
                else:
                    _on_download_complete(self.download_id, "failed")
                    
                _on_download_complete(self.download_id, "failed")
                log(f"Download failed for ID {self.download_id}: {stderr}")
        except FileNotFoundError:
            _on_download_complete(self.download_id, "failed")
            ui.message(_("Error Conversion not found"))
        except Exception as e:
            _on_download_complete(self.download_id, "failed")
            ui.message(_("Download error"))
            log(f"Download error: {str(e)}")
        finally:
            processID = None

    def _verify_final_file(self, expected_filepath_stem):
        search_pattern = f"{expected_filepath_stem}*.{self.format}"
        found_files = glob.glob(search_pattern)

        if not found_files:
            return False

        for filepath in found_files:
            try:
                if os.path.exists(filepath) and os.path.getsize(filepath) >= 1024:
                    return True
            except Exception:
                continue
        return False

    def _cleanup_temp_files(self, force=False):
        if not self.title or not self.path or not self.format:
            return

        base_filename = os.path.join(self.path, self.title)

        temp_patterns = [
            f"{base_filename}*.part",
            f"{base_filename}*.ytdl",
            f"{base_filename}*.temp",
            f"{base_filename}*.download",
            f"{base_filename}*.webm",
            f"{base_filename}*.m4a",
            f"{base_filename}*.f*.tmp",
            f"{base_filename}*.f*.mp4",
            os.path.join(self.path, "*.part"),
            os.path.join(self.path, "*.ytdl"),
            os.path.join(self.path, "*.temp"),
            os.path.join(self.path, "*.download"),
            os.path.join(self.path, "*.webm"),
            os.path.join(self.path, "*.m4a"),
            os.path.join(self.path, "*.f*.tmp"),
            os.path.join(self.path, "*.f*.mp4")
        ]

        for pattern in temp_patterns:
            for temp_file in glob.glob(pattern):
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                except Exception as e:
                    log(f"Error deleting temp file {temp_file}: {e}")

        if self.is_playlist or force:
            for i in range(100):
                numbered_patterns = [
                    f"{base_filename}{i:03d}.*",
                    os.path.join(self.path, f"*{i:03d}.*")
                ]
                for pattern in numbered_patterns:
                    for temp_file in glob.glob(pattern):
                        try:
                            if os.path.exists(temp_file):
                                os.remove(temp_file)
                        except Exception as e:
                            log(f"Error deleting temp file {temp_file}: {e}")


class WaitThread(threading.Thread):
    def __init__(self, targetThread):
        super().__init__()
        self.target = targetThread
        self.daemon = True

    def run(self):
        i = 0
        while self.target.is_alive():
            time.sleep(0.1)
            i += 1
            if i == 10:
                try:
                    PlayWave("heart")
                except Exception:
                    pass
                i = 0
        self.target.join()
