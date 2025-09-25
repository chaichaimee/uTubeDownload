# uTubeDownload_settings.py

import wx
import gui
import config
from gui.settingsDialogs import SettingsPanel
from gui import guiHelper

AddOnSummary = "uTubeDownload"
AddOnName = "uTubeDownload"
sectionName = AddOnName

def getINI(key):
    return config.conf[sectionName][key]

def setINI(key, value):
    config.conf[sectionName][key] = value

class AudioYoutubeDownloadPanel(SettingsPanel):
    title = AddOnSummary

    def makeSettings(self, settingsSizer):
        helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

        folderSizer = wx.StaticBoxSizer(wx.VERTICAL, self, label=_("Destination folder"))
        folderBox = folderSizer.GetStaticBox()
        folderHelper = guiHelper.BoxSizerHelper(self, sizer=folderSizer)

        browseText = _("&Browse...")
        dirDialogTitle = _("Select a directory")
        pathHelper = guiHelper.PathSelectionHelper(folderBox, browseText, dirDialogTitle)
        pathCtrl = folderHelper.addItem(pathHelper)
        self.folderPathCtrl = pathCtrl.pathControl
        
        current_result_folder = getINI("ResultFolder")
        if not current_result_folder:
            import os
            AppData = os.environ["APPDATA"]
            self.folderPathCtrl.SetValue(os.path.join(AppData, "uTubeDownload"))
        else:
            self.folderPathCtrl.SetValue(current_result_folder)

        helper.addItem(folderSizer)

        self.beepChk = helper.addItem(
            wx.CheckBox(self, label=_("&Beep while converting"))
        )
        self.beepChk.SetValue(getINI("BeepWhileConverting"))

        self.sayCompleteChk = helper.addItem(
            wx.CheckBox(self, label=_("&Say download complete"))
        )
        self.sayCompleteChk.SetValue(getINI("SayDownloadComplete"))

        qualityLabel = _("MP3 &quality (kbps):")
        self.qualityChoice = helper.addLabeledControl(
            qualityLabel,
            wx.Choice,
            choices=["320", "256", "192", "128"]
        )
        try:
            self.qualityChoice.SetSelection(
                ["320", "256", "192", "128"].index(str(getINI("MP3Quality")))
            )
        except ValueError:
            self.qualityChoice.SetSelection(0)

        self.multipartChk = helper.addItem(
            wx.CheckBox(self, label=_("Use download section"))
        )
        self.multipartChk.SetValue(getINI("UseMultiPart"))
        
        connectionsLabel = _("&Number of connections:")
        self.connectionsChoice = helper.addLabeledControl(
            connectionsLabel,
            wx.Choice,
            choices=[str(i) for i in range(1, 17)]
        )
        try:
            self.connectionsChoice.SetSelection(
                getINI("MultiPartConnections") - 1
            )
        except Exception:
            self.connectionsChoice.SetSelection(7)

        self.playlistModeChk = helper.addItem(
            wx.CheckBox(self, label=_("Enable &playlist mode by default"))
        )
        self.playlistModeChk.SetValue(getINI("PlaylistMode"))

        self.skipExistingChk = helper.addItem(
            wx.CheckBox(self, label=_("Skip existing files"))
        )
        self.skipExistingChk.SetValue(getINI("SkipExisting"))

        self.resumeOnRestartChk = helper.addItem(
            wx.CheckBox(self, label=_("Resume interrupted downloads on restart"))
        )
        self.resumeOnRestartChk.SetValue(getINI("ResumeOnRestart"))

        self.loggingChk = helper.addItem(
            wx.CheckBox(self, label=_("Enable &logging"))
        )
        self.loggingChk.SetValue(getINI("Logging"))

    def onSave(self):
        folder = self.folderPathCtrl.GetValue().strip()
        if folder.endswith("\\"):
            folder = folder[:-1]

        if folder:
            import os
            if not os.path.isdir(folder):
                try:
                    os.makedirs(folder, exist_ok=True)
                except Exception:
                    gui.messageBox(
                        _("Failed to create the specified folder. Please select a valid folder."),
                        _("Error"), wx.OK | wx.ICON_ERROR
                    )
                    return
            setINI("ResultFolder", folder)
            setINI("BeepWhileConverting", self.beepChk.GetValue())
            setINI("SayDownloadComplete", self.sayCompleteChk.GetValue())
            setINI("MP3Quality", int(self.qualityChoice.GetStringSelection()))
            setINI("PlaylistMode", self.playlistModeChk.GetValue())
            setINI("SkipExisting", self.skipExistingChk.GetValue())
            setINI("ResumeOnRestart", self.resumeOnRestartChk.GetValue())
            setINI("Logging", self.loggingChk.GetValue())
            setINI("UseMultiPart", self.multipartChk.GetValue())
            setINI("MultiPartConnections", int(self.connectionsChoice.GetStringSelection()))



