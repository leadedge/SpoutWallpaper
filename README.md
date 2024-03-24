# SpoutWallpaper
A system tray application for Spout Senders or video files as live desktop wallpaper or show the Bing daily wallpaper.

### Spout receiver
If a Spout sender is running, it will be displayed immediately.  
If not running, a sender will be detected as soon as it starts.  
For images or video, the senders are not displayed but can be selected.  
To select a sender :
* Find SpoutWallPaper in the TaskBar tray area
* Right mouse click to select the sender

### Video player
* Select "Video" from the menu and choose the video file.

### Image
* Select "Image" from the menu and choose the image file

### Daily wallpaper
* Select "Daily" from the menu.

### Slideshow
* Select "Slideshow" from the menu and choose image folder, slide duration and "random" if required

### "About" for details.

At program close there is an option to keep the new wallpaper or restore the original.

## FFmpeg
SpoutWallpaper uses FFmpeg to play videos. 
FFmpeg.exe and FFprobe.exe are required.

* Go to [https://github.com/GyanD/codexffmpeg/releases](https://github.com/GyanD/codexffmpeg/releases)
* Choose the latest "Essentials" build. e.g. " ffmpeg-6.1.1-essentials_build.zip " and download the file.
* Unzip the archive and copy "bin\FFmpeg.exe" and "bin\FFprobe.exe" to : "SpoutWallpaper\DATA\FFMPEG\"
