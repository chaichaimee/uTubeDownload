# __init__.py
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
        "Logging": "boolean(default=False)",
        "PlaylistMode": "boolean(default=False)",
        "SkipExisting": "boolean(default=True)",
        "ResumeOnRestart": "boolean(default=True)",
        "MaxConcurrentDownloads": "integer(default=1)"
    }
    config.conf.spec[sectionName] = confspec

initConfiguration()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = AddOnSummary

    def __init__(self):
        super().__init__()
        
        from .uTubeDownload_core import (
            initialize_folders, resumeInterruptedDownloads,
            convertToMP, getCurrentDocumentURL,
            DownloadPath, setINI, PlayWave
        )
        
        self.core_functions = {
            'initialize_folders': initialize_folders,
            'resumeInterruptedDownloads': resumeInterruptedDownloads,
            'convertToMP': convertToMP,
            'getCurrentDocumentURL': getCurrentDocumentURL,
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
