# ActivityTracker
Application to monitor computer activity

1. Download ffmpeg (https://ffmpeg.zeranoe.com/builds/) and copy to Program Files directory. Make sure the path in config.ini matches the location of ffmpeg.exe.
2. Copy MKL.exe and config.ini to directory of choice and add a shortcut to the program in the user's start up folder (i.e. C:\Users\<%user_name%>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup

config.ini
----------
capture_path          = (String)  The path to which the logs, images, and video files will be stored

path_to_ffmpeg        = (String)  The path to ffmpeg.exe

caption_interval      = (Integer) The interval (in milliseconds) at which to capture the window title caption used for sorting the logs

image_capture_interval= (Integer) The interval (in milliseconds) to capture images of the desktop

capture_video         = (Boolean) If set to 1, captures the desktop activity to video

capture_images        = (Boolean) If set to 1, captures images of the desktop at the interval specified by image_capture_interval
