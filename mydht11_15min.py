#1 hour interval
#request parameters and display starting message
pause = 900
pauseM = pause/60
print("\n")
print("Now Loading. Approximate time between readings is {0} seconds or {1} minutes...".format(pause,pauseM))

#import instructions for communicating with DHT22, keeping time and logging workbooks
import Adafruit_DHT
import time
import datetime
import openpyxl
import num2words
import subprocess
import zipfile
from datetime import date
from openpyxl import load_workbook
from num2words import num2words
from subprocess import call
from num2words import num2words
from subprocess import call
from zipfile import ZipFile

#set up text to speech
cmd_beg= 'espeak '
cmd_end= ' 2>/dev/null' # To dump the std errors to /dev/null
pauseM_speak=num2words(pauseM)
call([cmd_beg+pauseM_speak+cmd_end], shell=True)
call([cmd_beg+"Minute"+cmd_end], shell=True)
call([cmd_beg+"Gap"+cmd_end], shell=True)
call([cmd_beg+"between"+cmd_end], shell=True)
call([cmd_beg+"readings"+cmd_end], shell=True)

#provide the sensor used and the gpio pin used
DHT_SENSOR = Adafruit_DHT.DHT22
DHT_PIN = 4

#load workbook
wb = load_workbook ('/home/pi/WeatherLog.xlsx')
sheet = wb["Sheet1"]

#start loop
try:
	while True:
#read humidity and temperature every "pause" amount of time, then print it alongside time, read it aloud and save this to workbook
		humidity, temperature = Adafruit_DHT.read(DHT_SENSOR, DHT_PIN)
		if humidity is not None and temperature is not None:
			today = date.today()
			now = datetime.datetime.now().time()
			print("\n")
			print(today,now)
			print("Temp={0:0.1f}C Humidity={1:0.1f}%".format(temperature, humidity))
			row = (today, now, temperature, humidity)
			sheet.append(row)
			x=num2words(temperature)
			y=num2words(humidity)
			call([cmd_beg+x+cmd_end], shell=True)
			call([cmd_beg+"degrees"+cmd_end], shell=True)
			call([cmd_beg+y+cmd_end], shell=True)
			call([cmd_beg+"percent"+cmd_end], shell=True)
			call([cmd_beg+"humidity"+cmd_end], shell=True)
			time.sleep(0.1)
			wb.save('/home/pi/WeatherLog.xlsx')
			time.sleep(pause)
		else:
			print("Loading")
			time.sleep(2)
finally:
#when program is ended (with Ctrl+C) you need to save the workbook, it is also best to say bye to indicate when the save is complete
	print("\n Saving to workbook...")
	wb.save('/home/pi/WeatherLog.xlsx')
	time.sleep(1)
	print("\n Goodbye")
	time.sleep(1)


#next to add is it running at launch and shuts down after a reaasonable amount of time
##shutdown -h + 2
