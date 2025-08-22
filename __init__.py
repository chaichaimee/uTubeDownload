# __init__.py
# Copyright (C) 2025 ['chai chaimee]
# Licensed under GNU General Public License. See COPYING.txt for details.

import globalPluginHandler
from scriptHandler import script
import ui
import gui
import wx
from gui.settingsDialogs import NVDASettingsDialog
import config
import addonHandler
import os
import time
import datetime
import re
import glob

addonHandler.initTranslation()

AddOnSummary = "uTubeDownload"
AddOnName = "uTubeDownload"
AddOnPath = os.path.dirname(__file__)
sectionName = AddOnName
_last_tap_time = 0
_double_tap_threshold = 0.3

def initConfiguration():
    confspec = {
        "BeepWhileConverting": "boolean(default=True)",
        "ResultFolder": "string(default='')",
        "MP3Quality": "integer(default=320)",
        "TrimMP3Quality": "integer(default=320)",
        "Logging": "boolean(default=False)",
        "PlaylistMode": "boolean(default=False)",
        "SkipExisting": "boolean(default=True)",
        "ResumeOnRestart": "boolean(default=True)",
        "MaxConcurrentDownloads": "integer(default=1)",
        "TrimLastFormat": "string(default='mp3')",
        "TrimLastStartTime": "string(default='00:00:00')",
        "TrimLastEndTime": "string(default='00:00:00')",
        "UseMultiPart": "boolean(default=True)",
        "MultiPartConnections": "integer(default=8)",
    }
    config.conf.spec[sectionName] = confspec

initConfiguration()

def _find_next_trim_number(save_path):
    try:
        existing_files = glob.glob(os.path.join(save_path, "Trimmed Clip *.mp3"))
        existing_files.extend(glob.glob(os.path.join(save_path, "Trimmed Clip *.mp4")))
        numbers = []
        for file_path in existing_files:
            match = re.search(r"Trimmed Clip (\d+)\.(mp3|mp4)$", os.path.basename(file_path))
            if match:
                numbers.append(int(match.group(1)))
        if not numbers:
            return 1
        next_number = max(numbers) + 1
        return next_number
    except Exception:
        return 1

def _format_timedelta(seconds):
    """Convert seconds to HH:MM:SS format without days"""
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f"{int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}"

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = AddOnSummary

    def __init__(self):
        super().__init__()
        
        from .uTubeDownload_core import (
            initialize_folders, resumeInterruptedDownloads,
            convertToMP, getCurrentDocumentURL, getCurrentAppName,
            DownloadPath, setINI, PlayWave
        )
        
        self.core_functions = {
            'initialize_folders': initialize_folders,
            'resumeInterruptedDownloads': resumeInterruptedDownloads,
            'convertToMP': convertToMP,
            'getCurrentDocumentURL': getCurrentDocumentURL,
            'getCurrentAppName': getCurrentAppName,
            'DownloadPath': DownloadPath,
            'setINI': setINI,
            'PlayWave': PlayWave
        }
        
        self.core_functions['initialize_folders']()
        wx.CallAfter(self.core_functions['resumeInterruptedDownloads'])

        try:
            from .uTubeDownload_settings import AudioYoutubeDownloadPanel
            if AudioYoutubeDownloadPanel not in NVDASettingsDialog.categoryClasses:
                NVDASettingsDialog.categoryClasses.append(AudioYoutubeDownloadPanel)
        except ImportError:
            pass
        
        try:
            from .uTubeTrim import uTubeTrimDialog
        except ImportError:
            pass

    def terminate(self):
        try:
            from .uTubeDownload_settings import AudioYoutubeDownloadPanel
            if AudioYoutubeDownloadPanel in NVDASettingsDialog.categoryClasses:
                NVDASettingsDialog.categoryClasses.remove(AudioYoutubeDownloadPanel)
        except Exception:
            pass

    def _get_current_download_path(self):
        return config.conf[sectionName]["ResultFolder"] or self.core_functions['DownloadPath']

    @script(description=_("Download MP3 (single tap) or MP4 (double tap)"), gesture="kb:NVDA+y")
    def script_downloadMP3OrMP4(self, gesture):
        global _last_tap_time
        current_time = time.time()
        is_double_tap = (current_time - _last_tap_time) < _double_tap_threshold
        _last_tap_time = current_time
        
        if is_double_tap:
            url = self.core_functions['getCurrentDocumentURL']()
            if url:
                self.core_functions['PlayWave']('start')
                self.core_functions['convertToMP']("mp4", self._get_current_download_path(), config.conf[sectionName]["PlaylistMode"])
        else:
            def do_single_tap_action():
                if (time.time() - _last_tap_time) >= _double_tap_threshold:
                    url = self.core_functions['getCurrentDocumentURL']()
                    if url:
                        self.core_functions['PlayWave']('start')
                        self.core_functions['convertToMP']("mp3", self._get_current_download_path(), config.conf[sectionName]["PlaylistMode"])
            wx.CallLater(int(_double_tap_threshold * 1000), do_single_tap_action)

    @script(description=_("Open download folder"), gesture="kb:NVDA+control+y")
    def script_openDownloadFolder(self, gesture):
        path = self._get_current_download_path()
        if os.path.isdir(path):
            try:
                os.startfile(path)
            except Exception:
                ui.message(_("Error opening folder"))
        else:
            ui.message(_("Invalid download folder"))

    @script(description=_("Toggle playlist mode"), gesture="kb:NVDA+shift+y")
    def script_togglePlaylistMode(self, gesture):
        current_mode = config.conf[sectionName]["PlaylistMode"]
        self.core_functions['setINI']("PlaylistMode", not current_mode)
        ui.message(_("Playlist mode enabled") if not current_mode else _("Playlist mode disabled"))
    
    @script(description=_("uTubeTrim setting"), gesture="kb:NVDA+alt+y")
    def script_uTubeTrim(self, gesture):
        from .uTubeTrim import uTubeTrimDialog
        url = None
        for _ in range(3):
            url = self.core_functions['getCurrentDocumentURL']()
            if url:
                break
            time.sleep(0.5)
        
        def show_dialog():
            gui.mainFrame.prePopup()
            dlg = uTubeTrimDialog(gui.mainFrame, url or "")
            dlg.ShowModal()
            dlg.Destroy()
            gui.mainFrame.postPopup()
        
        wx.CallAfter(show_dialog)
    
    @script(description=_("uTubeSnapshot"), gesture="kb:control+shift+y")
    def script_captureSnapshot(self, gesture):
        # Only activate on YouTube URLs
        url = self.core_functions['getCurrentDocumentURL']()
        if not url:
            gesture.send()
            return
        
        url_lower = url.lower()
        if "youtube.com" not in url_lower and "youtu.be" not in url_lower:
            gesture.send()
            return

        from .uTubeSnapshot import capture_snapshot
        path = self._get_current_download_path()
        capture_snapshot(url, path)
