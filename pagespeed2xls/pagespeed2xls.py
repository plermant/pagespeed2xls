#! /usr/bin/env python
# by Pierre Lermant, January 2016
""" Copyright 2016 Akamai Technologies, Inc. All Rights Reserved.
 
 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at 
    http://www.apache.org/licenses/LICENSE-2.0
 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
"""

import sys, getopt
import urllib
import urllib2
import datetime
import time
import xlwt
import os
import json

from libs import printSpeed, printUsability, getJson
from time import gmtime, strftime
from sys import argv
#try ... except block needed for python 2/3 compatibility
try:
	from urllib2 import Request, urlopen, URLError, HTTPError
except ImportError:
	from urllib.request import urlopen, Request, URLError, HTTPError

#Global parameters ********************************************

#default file names
outFileName="pagespeed-"+str(time.time())+".xls" # make the default output file name unique, by appending the unix timestamp
inFileName=""#no input file by default
tmpJson="tmpJson"#temporary json files returned by google api are saved as 'tmpJson+"Desktop/Mobile"+str(i+1)+".json"' in case it's needed for troubleshooting

#google latest api path and optimized asset location
apiCall="https://www.googleapis.com/pagespeedonline/v2/runPagespeed?" 
googleStore="https://developers.google.com/speed/pagespeed/insights/optimizeContents?url="

#create overall Workbook
wb = xlwt.Workbook()

#spreadsheet styles

# add new colours to palette and set RGB colour values
#see styles at https://secure.simplistix.co.uk/svn/xlwt/tags/0.7.2/xlwt/Style.py
xlwt.add_palette_colour("custom_green", 0x21)
wb.set_colour_RGB(0x21, 200, 255, 200)
xlwt.add_palette_colour("custom_orange", 0x22)
wb.set_colour_RGB(0x22, 255, 250, 180)
xlwt.add_palette_colour("custom_red", 0x23)
wb.set_colour_RGB(0x23, 255, 200, 200)
xlwt.add_palette_colour("bright_red", 0x24)
wb.set_colour_RGB(0x24, 255, 0, 0)

#define cell styles
boldStyle=xlwt.Style.easyxf("font: bold on")
greenStyle=xlwt.easyxf('pattern: pattern solid,fore_colour custom_green;align: vert top')
orangeStyle=xlwt.easyxf('pattern: pattern solid,fore_colour custom_orange;align: vert top')
redStyle=xlwt.easyxf('pattern: pattern solid,fore_colour custom_red;align: vert top')
brightRedStyle=xlwt.easyxf('pattern: pattern solid,fore_colour bright_red;align: vert top')
defaultStyle=xlwt.easyxf('align: vert top')

styles={'defaultStyle':defaultStyle,'boldStyle':boldStyle,'greenStyle':greenStyle,'orangeStyle':orangeStyle,'redStyle':redStyle,'brightRedStyle':brightRedStyle}

#thresholds that dictate color coding in spreadsheet
scoreRedThreshold=65#red for this and under
scoreOrangeThreshold=85#green above this
impactRedThreshold=10# if 0 it's green, less than this number it's orange, more it's red

columnW=30#default column width
urlW=50#url column width
scoreW=6#score column width
dateW=10#date column width
		
#MAIN PROGRAM **************************************************************

#input parameters
if len(sys.argv) == 1:
	print("If you want to specify parameters through files, use "+sys.argv[0]+" -i <inputfile> -o <outputfile> and abort this program now ... \n"+"If you continue you'll be prompted for your api key and url to test.\n"+"Default output file name is "+outFileName)
	apiKey=raw_input("Please enter your apiKey: ")
	testUrl=raw_input("Please enter your url (e.g. http://www.example.com): ")
	runs=1
else:
	try:
		opts, args = getopt.getopt(sys.argv[1:],"hi:o:",["ifile=","ofile="])
	except getopt.GetoptError:
		print ("2 modes:")
		print(sys.argv[0]+" (you'll be asked for api key and url), or:")
		print sys.argv[0]+' -i <inputfile> -o <outputfile.xls> (input file has apikey as first line, followed by URLs, one per line)'
		sys.exit(2)
	for opt, arg in opts:
		if opt == '-h':
			print ("2 modes:")
			print(sys.argv[0]+" (you'll be asked for api key and url), or:")
			print sys.argv[0]+' -i <inputfile> -o <outputfile.xls> (input file has apikey as first line, followed by URLs, one per line)'
			sys.exit(0)
		elif opt in ("-i", "--ifile"):
			inFileName = arg
			print("If program behaves unexpectedly, check that the firt line of your file holds your api key and each subsequent line an url to test.\nNote that api keys have request number restrictions so only tests the url you need.")
		elif opt in ("-o", "--ofile"):
			outFileName = arg

#get information from input file, if relevant
if (inFileName!=""):
	with open(inFileName) as f:
		lines = f.read().splitlines()
	if len(lines) < 2:
		print("Expecting at least 2 lines in input file, aborting ...")
		sys.exit(2)
	apiKey=lines[0]
	runs=len(lines)-1

if (inFileName!=""):
	print('Input file is '+inFileName)
print('Output file is '+outFileName)

#Create excel sheets
wsMS=wb.add_sheet("Mobile Speed")
wsMU=wb.add_sheet("Mobile Usability")
wsDS=wb.add_sheet("Desktop Speed")

#Print sheet headers:
#Usability
ws=wsMU
#set default width to a set number of char
for i in range(0, 9):
	ws.col(i).width=columnW*256
ws.write(0,0,"URL and Title",boldStyle)
ws.col(0).width=urlW*256
ws.write(0,1,"Date GMT",boldStyle)
ws.col(1).width=dateW*256
ws.write(0,2,"Score",boldStyle)
ws.col(2).width=scoreW*256
ws.write(0,3,"Avoid Interstitials",boldStyle)
ws.write(0,4,"Avoid Plugins",boldStyle)
ws.write(0,5,"Configure Viewport",boldStyle)
ws.write(0,6,"Size Content To Viewport",boldStyle)
ws.write(0,7,"Size Tap Targets Appropriately",boldStyle)
ws.write(0,8,"Use Legible Font Sizes",boldStyle)
		
#Desktop and Mobile Speed
for j in range(0,2):
	if j==0:
		ws=wsDS
	if j==1:
		ws=wsMS
	#set default width to a set number of char
	for i in range(0, 13):
		ws.col(i).width=columnW*256
	ws.write(0,0,"URL and Title",boldStyle)
	ws.col(0).width=urlW*256
	ws.write(0,1,"Date GMT",boldStyle)
	ws.col(1).width=dateW*256
	ws.write(0,2,"Score",boldStyle)
	ws.col(2).width=scoreW*256
	ws.write(0,3,"Avoid Redirects",boldStyle)
	ws.write(0,4,"Enable Gzip",boldStyle)	
	ws.write(0,5,"Browser Caching",boldStyle)
	ws.write(0,6,"Main Response Time",boldStyle)
	ws.write(0,7,"Minify CSS",boldStyle)
	ws.write(0,8,"Minify HTML",boldStyle)
	ws.write(0,9,"Minify JS",boldStyle)
	ws.write(0,10,"Avoid Blocking Resources",boldStyle)
	ws.write(0,11,"Optimize Images",boldStyle)
	ws.write(0,12,"Prioritize Visible Content",boldStyle)
	ws.col(13).width=urlW*256
	ws.write(0,13,"Link to google storage of optimized resources",boldStyle)
		
		
#call api for each url for mobile and desktop experiences
try:
	if runs==0:
		print("Your input is empty")
except:
	print("Incorrect input, please see:")
	print sys.argv[0]+' -h'
	sys.exit(0)
for i in range(0,runs):
	if (inFileName!=""):
		testUrl=lines[i+1]
	if testUrl=="":
		print("Empty url line, skipping ...")
		continue
	print("Processing url "+testUrl+" please wait ...")
	try:
		#mobile
		apiCmd=apiCall+"url="+testUrl+"&strategy="+"mobile"+"&key="+apiKey
		jsonT=getJson(apiCmd)
		jsonO=json.loads(jsonT)
		#save output, in case ...
		tmpName=tmpJson+"Mobile"+str(i+1)+".json"
		if os.path.isfile(tmpName):
			os.remove(tmpName)
		_file = open(tmpName, "w")
		_file.write(jsonT)
		_file.close()
		printSpeed(wsMS,jsonO,1+i,styles,scoreOrangeThreshold,scoreRedThreshold,impactRedThreshold)
		printUsability(wsMU,jsonO,1+i,styles,scoreOrangeThreshold,scoreRedThreshold,impactRedThreshold)
		#Link to google storage of optimized resources
		body=googleStore+testUrl+"&strategy=mobile"
		wsMS.write(i+1,13,body)
		#desktop
		apiCmd=apiCall+"url="+testUrl+"&strategy="+"desktop"+"&key="+apiKey
		jsonT=getJson(apiCmd)
		jsonO=json.loads(jsonT)
		#save output, in case ...
		tmpName=tmpJson+"Desktop"+str(i+1)+".json"
		if os.path.isfile(tmpName):
			os.remove(tmpName)
		_file = open(tmpName, "w")
		_file.write(jsonT)
		_file.close()
		printSpeed(wsDS,jsonO,1+i,styles,scoreOrangeThreshold,scoreRedThreshold,impactRedThreshold)
		#Link to google storage of optimized resources
		body=googleStore+testUrl+"&strategy=desktop"
		wsDS.write(i+1,13,body)
	except:
		print("url FAILED, skipping ...")

#save Output
if os.path.isfile(outFileName):
	print("File "+outFileName+" already exists, either abort or rename output to something else")
	outFileName=raw_input("Please enter new output filename: ")
wb.save(outFileName)
