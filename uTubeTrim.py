# uTubeTrim.py

import wx
import gui
import config
from gui import guiHelper
import os
import urllib
import api
import ui
import threading
import subprocess
import re
import json
import uuid
import glob
import winsound
from .uTubeDownload_core import (
    addDownloadToQueue,
    getINI,
    log,
    YouTubeEXE,
    ConverterPath,
    _process_next_download,
    setINI,
    DownloadPath,
    _download_queue
)

AddOnName = "uTubeDownload"
sectionName = AddOnName

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

def _stop_all_sounds():
    try:
        winsound.PlaySound(None, winsound.SND_PURGE)
    except Exception:
        pass

class uTubeTrimDialog(wx.Dialog):
    def __init__(self, parent, initial_url=""):
        super().__init__(parent, title=_("uTubeTrim"), style=wx.DEFAULT_DIALOG_STYLE)
        self.download_path = getINI("ResultFolder") or DownloadPath
        self.parent = parent
        self.initial_url = initial_url
        self.video_duration = "00:00:00"
        self.video_duration_seconds = 0
        self.init_ui()
        self.urlCtrl.SetValue(self.initial_url)
        self.urlCtrl.SetFocus()
        self.startTimeCtrl.SetValue(config.conf[sectionName].get("TrimLastStartTime", "00:00:00"))
        
        # Use end time from config only for matching previous URL
        last_url = config.conf[sectionName].get("TrimLastURL", "")
        if initial_url and initial_url == last_url:
            end_time_value = config.conf[sectionName].get("TrimLastEndTime", "")
        else:
            end_time_value = ""
        self.endTimeCtrl.SetValue(end_time_value)
        
        wx.CallAfter(self.update_duration_label)
        last_format = config.conf[sectionName].get("TrimLastFormat", "mp3")
        last_quality = config.conf[sectionName].get("TrimLastQuality", getINI("TrimMP3Quality"))
        if last_format == "mp4":
            self.mp4Radio.SetValue(True)
        else:
            self.mp3Radio.SetValue(True)
        self.qualityCtrl.SetStringSelection(f"{last_quality} kbps")
        self.on_format_change(None)
        
        # Use cached duration for previous URL, fetch new for new URL
        if initial_url:
            last_url = config.conf[sectionName].get("TrimLastURL", "")
            last_duration = config.conf[sectionName].get("TrimLastDuration", "")
            if initial_url == last_url and last_duration:
                # Use cached duration
                self.video_duration = last_duration
                try:
                    parts = last_duration.split(':')
                    if len(parts) == 3:
                        hours, mins, secs = map(float, parts)
                        self.video_duration_seconds = hours*3600 + mins*60 + secs
                    elif len(parts) == 2:
                        mins, secs = map(float, parts)
                        self.video_duration_seconds = mins*60 + secs
                    else:
                        self.video_duration_seconds = float(parts[0])
                except ValueError:
                    log(f"Invalid cached duration: {last_duration}")
                    self.video_duration_seconds = 0
                wx.CallAfter(self.update_duration_label)
            else:
                # New URL: fetch video duration
                threading.Thread(target=self._fetch_video_duration, daemon=True).start()

    def init_ui(self):
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        urlSizer = wx.StaticBoxSizer(wx.VERTICAL, self, label=_("URL"))
        self.urlCtrl = wx.TextCtrl(self)
        urlSizer.Add(self.urlCtrl, 1, wx.EXPAND | wx.ALL, 5)
        mainSizer.Add(urlSizer, 0, wx.EXPAND | wx.ALL, 5)
        timeSizer = wx.BoxSizer(wx.VERTICAL)
        timeHelper = guiHelper.BoxSizerHelper(self, sizer=timeSizer)
        self.startTimeCtrl = timeHelper.addLabeledControl(_("Start time:"), wx.TextCtrl, value="00:00:00")
        self.endTimeCtrl = timeHelper.addLabeledControl(_("End time:"), wx.TextCtrl, value="00:00:00")
        mainSizer.Add(timeSizer, 0, wx.EXPAND | wx.ALL, 5)
        previewSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.previewDurationLabel = wx.StaticText(self, label=_("Video duration: unknown; Overall period: 0 seconds"))
        previewSizer.Add(self.previewDurationLabel, 1, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.previewBtn = wx.Button(self, label=_("Preview start"))
        self.previewBtn.Bind(wx.EVT_BUTTON, self.on_preview)
        previewSizer.Add(self.previewBtn, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        mainSizer.Add(previewSizer, 0, wx.EXPAND | wx.ALL, 5)
        formatSizer = wx.StaticBoxSizer(wx.VERTICAL, self, label=_("File Format"))
        self.mp3Radio = wx.RadioButton(self, label="MP3", style=wx.RB_GROUP)
        self.mp4Radio = wx.RadioButton(self, label="MP4")
        formatSizer.Add(self.mp3Radio, 0, wx.ALL, 5)
        formatSizer.Add(self.mp4Radio, 0, wx.ALL, 5)
        mainSizer.Add(formatSizer, 0, wx.EXPAND | wx.ALL, 5)
        qualitySizer = wx.StaticBoxSizer(wx.VERTICAL, self, label=_("MP3 Quality"))
        choices = ["320 kbps", "256 kbps", "192 kbps", "128 kbps"]
        self.qualityCtrl = wx.ComboBox(self, choices=choices, style=wx.CB_READONLY)
        self.qualityCtrl.SetStringSelection("320 kbps")
        qualitySizer.Add(self.qualityCtrl, 0, wx.ALL, 5)
        mainSizer.Add(qualitySizer, 0, wx.EXPAND | wx.ALL, 5)
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.downloadBtn = wx.Button(self, label=_("Start download"))
        self.downloadBtn.Bind(wx.EVT_BUTTON, self.on_download)
        btnSizer.Add(self.downloadBtn, 0, wx.ALL, 5)
        self.cancelBtn = wx.Button(self, label=_("Cancel"))
        self.cancelBtn.Bind(wx.EVT_BUTTON, self.on_cancel)
        btnSizer.Add(self.cancelBtn, 0, wx.ALL, 5)
        mainSizer.Add(btnSizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        self.SetSizerAndFit(mainSizer)
        self.mp3Radio.Bind(wx.EVT_RADIOBUTTON, self.on_format_change)
        self.mp4Radio.Bind(wx.EVT_RADIOBUTTON, self.on_format_change)
        self.startTimeCtrl.Bind(wx.EVT_TEXT, self.on_time_control_text)
        self.endTimeCtrl.Bind(wx.EVT_TEXT, self.on_time_control_text)
        self.on_format_change(None)

    def on_time_control_text(self, event):
        wx.CallAfter(self.update_duration_label)
        event.Skip()

    def _fetch_video_duration(self):
        url = self.urlCtrl.GetValue().strip()
        if not url:
            return
        try:
            cmd = [YouTubeEXE, "--get-duration", "--no-playlist", url]
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if result.returncode == 0:
                duration_str = result.stdout.strip()
                wx.CallAfter(self._update_duration, duration_str)
        except Exception as e:
            log(f"Error fetching video duration: {e}")

    def _update_duration(self, duration_str):
        if not self or not self.IsShown() or not duration_str:
            return
        self.video_duration = duration_str
        try:
            parts = duration_str.split(':')
            if len(parts) == 3:
                hours, mins, secs = map(float, parts)
                self.video_duration_seconds = hours*3600 + mins*60 + secs
            elif len(parts) == 2:
                mins, secs = map(float, parts)
                self.video_duration_seconds = mins*60 + secs
            else:
                self.video_duration_seconds = float(parts[0])
            # Save URL and duration for next time
            config.conf[sectionName]["TrimLastURL"] = self.initial_url
            config.conf[sectionName]["TrimLastDuration"] = duration_str
            # Set end time to video duration if not set
            if self.endTimeCtrl.GetValue() == "":
                self.endTimeCtrl.SetValue(duration_str)
        except ValueError:
            log(f"Invalid duration format: {duration_str}")
        wx.CallAfter(self.update_duration_label)

    def update_duration_label(self):
        if not self or not self.IsShown():
            return
        if self.video_duration_seconds > 0:
            video_dur_str = self.video_duration
        else:
            video_dur_str = _("unknown")
        try:
            start_seconds = self._time_str_to_seconds(self.startTimeCtrl.GetValue())
            end_seconds = self._time_str_to_seconds(self.endTimeCtrl.GetValue())
            period = end_seconds - start_seconds
            if period < 0:
                period_str = _("invalid (negative)")
            else:
                period_str = f"{period:.2f}"
        except (ValueError, IndexError):
            period_str = _("invalid")
        label = _("Video duration: {duration}; Overall period: {period} seconds").format(
            duration=video_dur_str,
            period=period_str
        )
        if self.previewDurationLabel:
            self.previewDurationLabel.SetLabel(label)

    def on_format_change(self, event):
        if self.qualityCtrl:
            self.qualityCtrl.Enable(self.mp3Radio.GetValue())

    def _time_str_to_seconds(self, time_str):
        parts = list(map(float, time_str.split(':')))
        if len(parts) == 3:
            return parts[0] * 3600 + parts[1] * 60 + parts[2]
        elif len(parts) == 2:
            return parts[0] * 60 + parts[1]
        elif len(parts) == 1:
            return parts[0]
        return 0

    def on_preview(self, event):
        start_time_str = self.startTimeCtrl.GetValue()
        start_seconds = self._time_str_to_seconds(start_time_str)
        if not self.initial_url or start_seconds < 0 or start_seconds >= self.video_duration_seconds:
            ui.message(_("Invalid URL or start time"))
            return
        preview_url = f"{self.initial_url}&t={int(start_seconds)}s"
        try:
            os.startfile(preview_url)
        except Exception:
            ui.message(_("Could not open preview URL"))

    def on_download(self, event):
        url = self.urlCtrl.GetValue().strip()
        if not url:
            ui.message(_("URL is required"))
            return
        
        start_time_str = self.startTimeCtrl.GetValue()
        end_time_str = self.endTimeCtrl.GetValue()
        start_seconds = self._time_str_to_seconds(start_time_str)
        end_seconds = self._time_str_to_seconds(end_time_str)
        
        if start_seconds >= end_seconds:
            ui.message(_("Start time must be less than end time"))
            return

        file_format = "mp3" if self.mp3Radio.GetValue() else "mp4"
        quality_kbps = int(self.qualityCtrl.GetStringSelection().split()[0])
        trim_number = _find_next_trim_number(self.download_path)
        output_filename = f"Trimmed Clip {trim_number}"
        output_path = os.path.join(self.download_path, output_filename)

        if not os.path.exists(YouTubeEXE):
            ui.message(_("Error: yt-dlp.exe not found."))
            return
        if not os.path.exists(os.path.join(ConverterPath, "ffmpeg.exe")):
            ui.message(_("Error: ffmpeg.exe not found."))
            return
        
        ui.message(_("Trimming and downloading {format}...").format(format=file_format.upper()))
        
        download_sections_arg = f"*{start_time_str}-{end_time_str}"
        
        base_cmd = [
            YouTubeEXE,
            url,
            "--no-playlist",
            "-o", f"{output_path}.%(ext)s",
            "--ffmpeg-location", os.path.join(ConverterPath, "ffmpeg.exe"),
            "--download-sections", download_sections_arg
        ]

        if file_format == "mp3":
            base_cmd.extend([
                "--extract-audio",
                "--audio-format", "mp3",
                "--audio-quality", str(quality_kbps)
            ])
        else: # mp4
            base_cmd.extend([
                "-f", "bestvideo+bestaudio/best",
                "--merge-output-format", "mp4",
                "--postprocessor-args", "ffmpeg:-c:v libx265 -preset fast -crf 23 -c:a copy"
            ])
            # Note: The trimming is handled by --download-sections,
            # but the postprocessor args for video encoding are still needed.

        download_obj = {
            "url": url,
            "title": output_filename,
            "format": file_format,
            "path": self.download_path,
            "cmd": base_cmd,
            "is_playlist": False,
            "trimming": True
        }
        download_id = addDownloadToQueue(download_obj)
        download_obj["id"] = download_id
        _download_queue.put(download_obj)

        config.conf[sectionName]["TrimLastFormat"] = file_format
        config.conf[sectionName]["TrimLastQuality"] = quality_kbps
        config.conf[sectionName]["TrimLastStartTime"] = start_time_str
        config.conf[sectionName]["TrimLastEndTime"] = end_time_str
        self.Close()
    
    def on_cancel(self, event):
        self.Close()

