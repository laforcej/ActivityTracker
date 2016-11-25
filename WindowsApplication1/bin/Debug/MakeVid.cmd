echo off
set arg1=%1
cd "C:\Program Files\ffmpeg\bin\"
ffmpeg -framerate 1/1 -i %arg1%\img%%05d.png -c:v libx264 -r 30 -pix_fmt yuv420p %arg1%\out.mp4