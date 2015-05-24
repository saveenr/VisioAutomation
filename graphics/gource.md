
# Getting the tools

## Gource

First install the gource tool from here: https://github.com/acaudwell/Gource/releases

For example store it here:
	
	E:\gource-0.42.win32

## FFMPEG

Get the windows version from here: http://ffmpeg.zeranoe.com/builds/

keep the files in here

	E:\ffmpeg


## Running Gource



	set path=%path%;C:\Program Files (x86)\Git\bin


	gource.exe E:\code\github\VisioAutomation --seconds-per-day 0.005 --title VisioAutomation --hide filenames,usernames --background 5555dd -viewport 1920x1080 -o d:\visioautomation.ppm

## Running ffmpeg

NOTE  opens correctly with VLC but not with WMP x264 co

	ffmpeg.exe -y -r 60 -f image2pipe -vcodec ppm -i d:\visioautomation.ppm -vcodec libx264 -preset ultrafast -crf 1 -threads 0 -bf 0 d:\visioautomation.mp4