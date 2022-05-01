from pymediainfo import MediaInfo
from tkinter import filedialog
import xlsxwriter
import os
import mimetypes
import datetime
from io import BytesIO
row = 0
col = 0
videoBool = "false"
def media(f, writer):
	try:
		media_info = MediaInfo.parse(f)
		FPS=""
		height=""
		width=""
		videoDataRate=""
		audioDataRate=""
		audioSamplingRate=""
		fileName=""
		fileExtension=""
		filePath=""
		fileSize=""
		duration=""
		taggeddate=""
		filecreationdate=""
		for track in media_info.tracks:
			if track.track_type == "Video":
				height = track.height
				width = track.width
				if track.bit_rate:
					videoDataRate= track.bit_rate/1000
				FPS=track.frame_rate
			elif track.track_type == "Audio":
				if track.bit_rate:
					audioDataRate = track.bit_rate/1000
				if track.sampling_rate:
					audioSamplingRate = track.sampling_rate/1000
			elif track.track_type == "General":
				fileName=track.file_name
				fileExtension=track.file_extension
				filePath = f
				if track.file_size:
					fileSize = track.file_size/1048576
				if track.other_duration[3]:
					duration = track.other_duration[3].split(".", 1)[0]
				taggeddate=track.tagged_date
				filecreationdate=track.file_creation_date
		content =[fileName,fileExtension,filePath,fileSize,duration,FPS, height, width, videoDataRate, audioDataRate, audioSamplingRate, taggeddate, filecreationdate]
		global row
		global col
		row += 1
		col = 0
		for item in content :
		    worksheet.write(row, col, item)
		    col += 1
	except Exception as e:
		print("Error occured in " , f, " : ", e)

def processFile(directory, filename, writer):
	global videoBool
	f = os.path.join(directory, filename)
	if (os.path.isfile(f) and mimetypes.guess_type(filename)[0] is not None and (mimetypes.guess_type(filename)[0].startswith("video"))):
		print(f,os.path.isfile(f), mimetypes.guess_type(filename)[0])
	# checking if it is a file
	if (os.path.isfile(f) and mimetypes.guess_type(filename)[0] is not None and (mimetypes.guess_type(filename)[0].startswith("video"))):
		print(f,os.path.isfile(f))				
		media(f, writer)
		videoBool = "true"

def loopD(directory, writer):
# iterate over files in
# that directory
	global videoBool
	for root, dirs, files in os.walk(directory):
		for filename in files:
			processFile(directory, filename, writer)
		for dir in dirs:
			for root2, dirs2, files2 in os.walk(os.path.join(root, dir)):
				for filename2 in files2:
					processFile(os.path.join(root, dir), filename2, writer)

try:
	directory = filedialog.askdirectory()
	workbook = xlsxwriter.Workbook(directory +'/VideoDetails_new.xlsx')
	worksheet = workbook.add_worksheet()
	content = ["File Name","File Extension","Path","Size in MB","Duration","Frame per second","Height","Width","Video Data rate kbps", "Audio Bit rate kbps","Audio Sample Rate kHz", "Tagged Date", "Creation Date"]
	for item in content :
		# write operation perform
		worksheet.write(row, col, item)
		col += 1
	loopD(directory, worksheet)
	if (videoBool == "true"):
		print("Writing to file...", directory +'/VideoDetails_new.xlsx')
	else:
		print("No video is present in the folder")
	key_pressed = input('Press ENTER to continue: ')
	workbook.close()
except Exception as e:
	print("Error occured: ", e)
	key_pressed = input('Press ENTER to continue: ')