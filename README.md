# uTubeDownload

**Author:** chai chaimee  

**URL:** [https://github.com/chaichaimee/uTubeDownload](https://github.com/chaichaimee/uTubeDownload)  


An add-on that allows you to download audio, video, and image files from [YouTube.com](https://www.youtube.com) with features supporting queue, resume, and download section.  


## Features  


**• uTubeDownload**  

**1.** Can choose to download as MP3 or MP4, either as a single file or playlist.  

**2.** Comes with a background system to ensure efficient downloading  

**2.1 Queue Manager**: Arranges downloads one by one to prevent heavy CPU and RAM usage (supports downloading a large number of files at once smoothly).  

**2.2 Resume System**: Supports continuous downloading after interruptions caused by restarting NVDA or turning Windows off/on (can download the remaining files on the next day or when NVDA is restarted).  
The resume window will prompt: if you choose **Yes**, it will resume from pending files; if you choose **No**, it will delete the remaining download history.  

**2.3 Automatic File Repair**: Repairs corrupted files caused by incomplete downloads automatically.  

**2.4 Downloaded File Check**: Skips files that are already downloaded to prevent duplication.  

**2.5 Download Section multi-part**: Can split the file into up to 16 parts to increase download speed by 50%.  

**2.6** All the above background systems can be enabled or disabled from NVDA Settings / uTubeDownload.  


**• uTubeTrim**  

Feature for trimming YouTube videos by setting start and end times (specific segment of the clip).  
Can select MP3 format with “Quality 128–320 kbps” and MP4 H.265.  
Includes a preview button to check the chosen start point before downloading.  
When on the YouTube page you want to trim, press `NVDA+ALT+Y` to open the uTubeTrim setting window.  


**• uTubeSnapshot**  

Can capture still images from a YouTube video into `.jpg` format using the shortcut `Control+Shift+Y`.  


## Keyboard Shortcuts  

- `NVDA+Y` - Download as MP3 (single tap)  
- `NVDA+Y` twice - Download as MP4 (double tap)  
- `NVDA+Shift+Y` - Toggle playlist mode  
- `NVDA+Ctrl+Y` - Open downloads folder  
- `NVDA+ALT+Y` - Open uTubeTrim Setting  
- `Control+Shift+Y` - Auto Snapshot  

All keyboard shortcuts can be changed in Input Gestures.  


## Donation  

If you like my work, you can donate to me via:  
[https://github.com/chaichaimee](https://github.com/chaichaimee)
