# Modified by Filip Pranklin, 2018.
# it's more or less not vulnerable to app shutting down (in terms of data loss) or deleting the worksheet being used (in terms of crashing)

# the added modification puts each days data in a separate workspace named dd/mm/yy in 3 columns



#!/usr/bin/python

# Google Spreadsheet DHT Sensor Data-logging Example

# Depends on the 'gspread' and 'oauth2client' package being installed.  If you
# have pip installed execute:
#   sudo pip install gspread oauth2client

# Also it's _very important_ on the Raspberry Pi to install the python-openssl
# package because the version of Python is a bit old and can fail with Google's
# new OAuth2 based authentication.  Run the following command to install the
# the package:
#   sudo apt-get update
#   sudo apt-get install python-openssl

# Copyright (c) 2014 Adafruit Industries
# Author: Tony DiCola

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
import json
import sys
import time
import datetime

import Adafruit_DHT
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Type of sensor, can be Adafruit_DHT.DHT11, Adafruit_DHT.DHT22, or Adafruit_DHT.AM2302.
DHT_TYPE = Adafruit_DHT.DHT22

# Example of sensor connected to Raspberry Pi pin 4
DHT_PIN  = 4
# Example of sensor connected to Beaglebone Black pin P8_11
#DHT_PIN  = 'P8_11'

# Google Docs OAuth credential JSON file.  Note that the process for authenticating
# with Google docs has changed as of ~April 2015.  You _must_ use OAuth2 to log
# in and authenticate with the gspread library.  Unfortunately this process is much
# more complicated than the old process.  You _must_ carefully follow the steps on
# this page to create a new OAuth service in your Google developer console:
#   http://gspread.readthedocs.org/en/latest/oauth2.html
#
# Once you've followed the steps above you should have downloaded a .json file with
# your OAuth2 credentials.  This file has a name like SpreadsheetData-<gibberish>.json.
# Place that file in the same directory as this python script.
#
# Now one last _very important_ step before updating the spreadsheet will work.
# Go to your spreadsheet in Google Spreadsheet and share it to the email address
# inside the 'client_email' setting in the SpreadsheetData-*.json file.  For example
# if the client_email setting inside the .json file has an email address like:
#   149345334675-md0qff5f0kib41meu20f7d1habos3qcu@developer.gserviceaccount.com
# Then use the File -> Share... command in the spreadsheet to share it with read
# and write acess to the email address above.  If you don't do this step then the
# updates to the sheet will fail!
GDOCS_OAUTH_JSON       = 'your SpreadsheetData-*.json file name'

# Google Docs spreadsheet name.
GDOCS_SPREADSHEET_NAME = 'tour google docs spreadsheet name'

# How long to wait (in seconds) between measurements.
FREQUENCY_SECONDS    = 60
MINIMUM_SENSOR_SLEEP = 2
NUMBER_OF_ROWS       = 1450     #one day is supposed to have 1440 entries with 60 second update frequency, a few extra rows just to be sure
NUMBER_OF_COLS       = 3        #Time / Temperature / Humidity

#Global variables for positioning inside of a worksheet
row = 1
col = 1

def get_spreadsheet(oauth_key_file, spreadsheet_name):
    try:
        scope =  ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(oauth_key_file, scope)
        gc = gspread.authorize(credentials)
        sh = gc.open(spreadsheet_name)
        return sh
    except Exception as ex:
        print('Unable to login and get spreadsheet.  Check OAuth credentials, spreadsheet name, and make sure spreadsheet is shared to the client_email address in the OAuth .json file!')
        print('Google sheet login failed with error:', ex)
        sys.exit(1)

def get_next_worksheet(spreadsheet, ttl, rws, cls):
    wh = None
    global row
    
    #Open the worksheet with desired title if it exists
    try:
        wh = sh.worksheet(ttl)
        print ("Opened worksheet " + ttl)
        #Find the row with the last enrty
        col_list = wh.col_values(2)
        row = 1
        for a in col_list:
            row = row + 1
            
    except gspread.exceptions.WorksheetNotFound as exc:
        #Create new worksheet with desired title otherwise
        print ("A worksheet with this name doesn't exist, creating one")
        wh = sh.add_worksheet(title = ttl, rows = rws, cols = cls)
        wh.update_cell(1, 1, "Time")
        wh.update_cell(1, 2, "Temperature")
        wh.update_cell(1, 3, "Humidity")
        row = 2
        
    except Exception as exc:
        print('Failed while choosing/creating workspace:', exc)
        
    return wh

print('Logging sensor measurements to {0} every {1} seconds.'.format(GDOCS_SPREADSHEET_NAME, FREQUENCY_SECONDS))
print('Press Ctrl-C to quit.')

sh = None
while True:
    global row
    global col

    # Login if necessary.
    if sh is None:
        sh = get_spreadsheet(GDOCS_OAUTH_JSON, GDOCS_SPREADSHEET_NAME)
       
    #Get worksheet with title equal to todays date and good row value
    wh = get_next_worksheet(sh, time.strftime("%d/%m/%Y"), NUMBER_OF_ROWS, NUMBER_OF_COLS)

    # Attempt to get sensor reading.
    humidity, temp = Adafruit_DHT.read(DHT_TYPE, DHT_PIN)

    # Skip to the next reading if a valid measurement couldn't be taken.
    # This might happen if the CPU is under a lot of load and the sensor
    # can't be reliably read (timing is critical to read the sensor).
    if humidity is None or temp is None:
        time.sleep(MINIMUM_SENSOR_SLEEP)
        continue
        
    print('Time:        ' + time.strftime("%H:%M:%S"))
    print('Temperature: {0:0.1f} C'.format(temp))
    print('Humidity:    {0:0.1f} %'.format(humidity))

    # Append the data in the worksheet, including a timestamp
    try:
        wh.update_cell(row, col, time.strftime("%H:%M:%S"))
        wh.update_cell(row, col+1, round(temp,1))
        wh.update_cell(row, col+2, round(humidity,1))
            
    except Exception as ex:
        # Error appending data, most likely because credentials are stale.
        # Null out the spreadsheet so a login is performed at the top of the loop.
        print(ex)
        print('Append error, logging in again')
        sh = None
        time.sleep(FREQUENCY_SECONDS)
        continue

    # Wait 30 seconds before continuing
    print('Wrote a row to {0} in {1}'.format(GDOCS_SPREADSHEET_NAME, wh))
    print("--------------------------------------------------------------")
    time.sleep(FREQUENCY_SECONDS)
