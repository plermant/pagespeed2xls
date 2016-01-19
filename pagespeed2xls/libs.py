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

from time import gmtime, strftime
from sys import argv
#try ... except block needed for python 2/3 compatibility
try:
	from urllib2 import Request, urlopen, URLError, HTTPError
except ImportError:
	from urllib.request import urlopen, Request, URLError, HTTPError

#these are strings provided by Google that help differentiate the kind of remediations to take.
expirationNotSpecified="{{URL}} (expiration not specified)"#unlike Google UI, I ignore short time caching, only highlighting resources not cached at all
losslessOnly="Losslessly compressing {{URL}} could save {{SIZE_IN_BYTES}} ({{PERCENTAGE}} reduction)."#compression is lossless
resize="Compress and resize {{URL}} could save {{SIZE_IN_BYTES}} ({{PERCENTAGE}} reduction)."#compression and resizing

#Subroutines **************************************************

def printSpeed(ws,jsonO,index,styles,scoreOrangeThreshold,scoreRedThreshold,impactRedThreshold):#print data into spreadsheet 'ws'. jsonO is the json object, result from api call. index is the index of the URL in the table, starting at 1, styles are cell styles
	defaultStyle=styles['defaultStyle']
	boldStyle=styles['boldStyle']
	greenStyle=styles['greenStyle']
	orangeStyle=styles['orangeStyle']
	redStyle=styles['redStyle']
	brightRedStyle=styles['brightRedStyle']
	
	#general data: url, title, score
	ws.write(index,1,strftime("%Y-%m-%d %H:%M:%S", gmtime()),defaultStyle)
	ws.write(index,0,jsonO['id']+' \x0a'+jsonO['title'],defaultStyle)
	if jsonO['ruleGroups']['SPEED']['score'] < scoreRedThreshold:
		style = redStyle
	elif jsonO['ruleGroups']['SPEED']['score'] < scoreOrangeThreshold:
		style = orangeStyle
	else:
		style = greenStyle
	ws.write(index,2,jsonO['ruleGroups']['SPEED']['score'],style)
		
	#AvoidLandingPageRedirects
	if (jsonO['formattedResults']['ruleResults']['AvoidLandingPageRedirects']['ruleImpact'] == 0):
		ws.write(index,3,'N/A',greenStyle)
	else:
		#rewrite the original URL, as the id is now the redirected url
		originalUrl=jsonO['formattedResults']['ruleResults']['AvoidLandingPageRedirects']['urlBlocks'][0]['urls'][0]['result']['args'][0]['value']
		numRedirects=jsonO['formattedResults']['ruleResults']['AvoidLandingPageRedirects']['summary']['args'][0]['value']
		strRedirects=""
		if(jsonO['formattedResults']['ruleResults']['AvoidLandingPageRedirects']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		for i in range(0,int(numRedirects)):
			strRedirects=strRedirects+jsonO['formattedResults']['ruleResults']['AvoidLandingPageRedirects']['urlBlocks'][0]['urls'][i+1]['result']['args'][0]['value']+'\x0a'
		body='This page has '+str(numRedirects)+" redirects:"+'\x0a'+"original URL is: "+'\x0a'+originalUrl+'\x0a'+"Redirected to: "+'\x0a'+strRedirects
		ws.write(index,3,body,style)
	
		
	#EnableGzipCompression
	if (jsonO['formattedResults']['ruleResults']['EnableGzipCompression']['ruleImpact'] == 0):
		ws.write(index,4,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['EnableGzipCompression']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		sizeReduction=jsonO['formattedResults']['ruleResults']['EnableGzipCompression']['urlBlocks'][0]['header']['args'][1]['value']
		sizePercent=jsonO['formattedResults']['ruleResults']['EnableGzipCompression']['urlBlocks'][0]['header']['args'][2]['value']
		body="Compress text files can save "+sizeReduction+", or "+sizePercent
		body=body+'\x0a'+"List of flies that can be compressed (savings):"
		urls=jsonO['formattedResults']['ruleResults']['EnableGzipCompression']['urlBlocks'][0]['urls']
		for url in urls:
			body=body+'\x0a'+url['result']['args'][0]['value'] + " ("+url['result']['args'][1]['value']+")"
		ws.write(index,4,body,style)
		
	#LeverageBrowserCaching
	if (jsonO['formattedResults']['ruleResults']['LeverageBrowserCaching']['ruleImpact'] == 0):
		ws.write(index,5,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['LeverageBrowserCaching']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		body="Consider caching on client these files: "
		urls=jsonO['formattedResults']['ruleResults']['LeverageBrowserCaching']['urlBlocks'][0]['urls']
		for url in urls:
			if url['result']['format'] == expirationNotSpecified:#ignore short time caching, only highlight resources not cached at all
				body=body+'\x0a'+url['result']['args'][0]['value']
		ws.write(index,5,body,style)
	
	#MainResourceServerResponseTime
	if (jsonO['formattedResults']['ruleResults']['MainResourceServerResponseTime']['ruleImpact'] == 0):
		ws.write(index,6,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['MainResourceServerResponseTime']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		body="Server Time is: "+str(jsonO['formattedResults']['ruleResults']['MainResourceServerResponseTime']['urlBlocks'][0]['header']['args'][0]['value'])
		ws.write(index,6,body,style)
		
	#MinifyCss
	if (jsonO['formattedResults']['ruleResults']['MinifyCss']['ruleImpact'] == 0):
		ws.write(index,7,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['MinifyCss']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		sizeReduction=jsonO['formattedResults']['ruleResults']['MinifyCss']['urlBlocks'][0]['header']['args'][1]['value']
		sizePercent=jsonO['formattedResults']['ruleResults']['MinifyCss']['urlBlocks'][0]['header']['args'][2]['value']
		body="Minifying CSS files can save "+sizeReduction+", or "+sizePercent
		body=body+'\x0a'+"List of CSS files that can be minimized (savings):"
		urls=jsonO['formattedResults']['ruleResults']['MinifyCss']['urlBlocks'][0]['urls']
		for url in urls:
			body=body+'\x0a'+url['result']['args'][0]['value'] + " ("+url['result']['args'][1]['value']+")"
		ws.write(index,7,body,style)
		
	#MinifyHTML
	if (jsonO['formattedResults']['ruleResults']['MinifyHTML']['ruleImpact'] == 0):
		ws.write(index,8,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['MinifyHTML']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		sizeReduction=jsonO['formattedResults']['ruleResults']['MinifyHTML']['urlBlocks'][0]['header']['args'][1]['value']
		sizePercent=jsonO['formattedResults']['ruleResults']['MinifyHTML']['urlBlocks'][0]['header']['args'][2]['value']
		body="Minifying HTML files can save "+sizeReduction+", or "+sizePercent
		body=body+'\x0a'+"List of HTML files that can be minimized (savings):"
		urls=jsonO['formattedResults']['ruleResults']['MinifyHTML']['urlBlocks'][0]['urls']
		for url in urls:
			body=body+'\x0a'+url['result']['args'][0]['value'] + " ("+url['result']['args'][1]['value']+")"
		ws.write(index,8,body,style)
		
	#MinifyJavaScript
	if (jsonO['formattedResults']['ruleResults']['MinifyJavaScript']['ruleImpact'] == 0):
		ws.write(index,9,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['MinifyJavaScript']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		sizeReduction=jsonO['formattedResults']['ruleResults']['MinifyJavaScript']['urlBlocks'][0]['header']['args'][1]['value']
		sizePercent=jsonO['formattedResults']['ruleResults']['MinifyJavaScript']['urlBlocks'][0]['header']['args'][2]['value']
		body="Minifying JS files can save "+sizeReduction+", or "+sizePercent
		body=body+'\x0a'+"List of JS files that can be minimized (savings):"
		urls=jsonO['formattedResults']['ruleResults']['MinifyJavaScript']['urlBlocks'][0]['urls']
		for url in urls:
			body=body+'\x0a'+url['result']['args'][0]['value'] + " ("+url['result']['args'][1]['value']+")"
		ws.write(index,9,body,style)
		
	#MinimizeRenderBlockingResources
	if (jsonO['formattedResults']['ruleResults']['MinimizeRenderBlockingResources']['ruleImpact'] == 0):
		ws.write(index,10,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['MinimizeRenderBlockingResources']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		body="List of files that may block rendering:"
		urls=jsonO['formattedResults']['ruleResults']['MinimizeRenderBlockingResources']['urlBlocks'][1]['urls']
		for url in urls:
			body=body+'\x0a'+url['result']['args'][0]['value']
		ws.write(index,10,body,style)
		
	#OptimizeImages
	if (jsonO['formattedResults']['ruleResults']['OptimizeImages']['ruleImpact'] == 0):
		ws.write(index,11,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['OptimizeImages']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		sizeReduction=jsonO['formattedResults']['ruleResults']['OptimizeImages']['urlBlocks'][0]['header']['args'][1]['value']
		sizePercent=jsonO['formattedResults']['ruleResults']['OptimizeImages']['urlBlocks'][0]['header']['args'][2]['value']
		body="Compress and resize can save "+sizeReduction+", or "+sizePercent
		body=body+'\x0a'+"List of flies that can be optimized (savings):"
		urls=jsonO['formattedResults']['ruleResults']['OptimizeImages']['urlBlocks'][0]['urls']
		start=0
		for url in urls:
			if url['result']['format'] == losslessOnly:
				if start == 0:
					body=body+'\x0a'+"Lossless compression only:"
					start=1
				body=body+'\x0a'+url['result']['args'][0]['value']+ " ("+url['result']['args'][1]['value']+")"
		start=0
		for url in urls:
			if url['result']['format'] == resize:
				if start == 0:
					body=body+'\x0a'+"Compress and resize:"
					start=1
				body=body+'\x0a'+url['result']['args'][0]['value']+ " ("+url['result']['args'][1]['value']+")"
		ws.write(index,11,body,style)
		
	#PrioritizeVisibleContent
	if (jsonO['formattedResults']['ruleResults']['PrioritizeVisibleContent']['ruleImpact'] == 0):
		ws.write(index,12,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['PrioritizeVisibleContent']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		body="Only "+jsonO['formattedResults']['ruleResults']['PrioritizeVisibleContent']['urlBlocks'][0]['urls'][0]['result']['args'][0]['value']+" of the final above-the-fold content rendered with the full HTML response"
		ws.write(index,12,body,style)
		
		
def printUsability(ws,jsonO,index,styles,scoreOrangeThreshold,scoreRedThreshold,impactRedThreshold):#print data into spreadsheet 'ws'. jsonO is the json object, result from api call. index is the index of the URL in the table, starting at 1	
	defaultStyle=styles['defaultStyle']
	boldStyle=styles['boldStyle']
	greenStyle=styles['greenStyle']
	orangeStyle=styles['orangeStyle']
	redStyle=styles['redStyle']
	brightRedStyle=styles['brightRedStyle']
	
	#general data: url, title, score
	ws.write(index,1,strftime("%Y-%m-%d %H:%M:%S", gmtime()),defaultStyle)
	ws.write(index,0,jsonO['id']+' \x0a'+jsonO['title'],defaultStyle)
	if jsonO['ruleGroups']['USABILITY']['score'] < scoreRedThreshold:
		style = redStyle
	elif jsonO['ruleGroups']['USABILITY']['score'] < scoreOrangeThreshold:
		style = orangeStyle
	else:
		style = greenStyle
	ws.write(index,2,jsonO['ruleGroups']['USABILITY']['score'],style)
		
	#Interstitial
	if (jsonO['formattedResults']['ruleResults']['AvoidInterstitials']['ruleImpact'] == 0):
		ws.write(index,3,'N/A',greenStyle)
	else:
		ws.write(index,3,'TBD, see https://developers.google.com/webmasters/mobile-sites/mobile-seo/common-mistakes/avoid-interstitials',brightRedStyle)

	#AvoidPlugins
	if (jsonO['formattedResults']['ruleResults']['AvoidPlugins']['ruleImpact'] == 0):
		ws.write(index,4,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['AvoidPlugins']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		body='Avoid Flash, Java, Silverlight or other plugins: '
		for url in jsonO['formattedResults']['ruleResults']['AvoidPlugins']['urlBlocks'][0]['urls']:
			body=body+'\x0a'+url['result']['args'][0]['value']
		ws.write(index,4,body,style)
		
	#ConfigureViewport
	if (jsonO['formattedResults']['ruleResults']['ConfigureViewport']['ruleImpact'] == 0):
		ws.write(index,5,'N/A',greenStyle)
	else:
		ws.write(index,5,'Your page does not have a viewport specified. This causes mobile devices to render your page as it would appear on a desktop browser, scaling it down to fit on a mobile screen',redStyle)
		
	#SizeContentToViewport
	if (jsonO['formattedResults']['ruleResults']['SizeContentToViewport']['ruleImpact'] == 0):
		ws.write(index,6,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['SizeContentToViewport']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		cWidth=jsonO['formattedResults']['ruleResults']['SizeContentToViewport']['urlBlocks'][0]['header']['args'][0]['value']
		vWidth=jsonO['formattedResults']['ruleResults']['SizeContentToViewport']['urlBlocks'][0]['header']['args'][1]['value']
		body='The page content is too wide for the viewport, forcing the user to scroll horizontally'+'\x0a'+'The page content is: '+cWidth+" CSS pixels wide, but the viewport is only: "+vWidth+" CSS pixels wide."+'\x0a'+"The following elements fall outside the viewport:"+'\x0a'
		for url in jsonO['formattedResults']['ruleResults']['SizeContentToViewport']['urlBlocks'][0]['urls']:
			body=body+url['result']['args'][0]['value']+'\x0a'
		ws.write(index,6,body,style)
		
	#SizeTapTargetsAppropriately
	if (jsonO['formattedResults']['ruleResults']['SizeTapTargetsAppropriately']['ruleImpact'] == 0):
		ws.write(index,7,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['SizeTapTargetsAppropriately']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		target=jsonO['formattedResults']['ruleResults']['SizeTapTargetsAppropriately']['urlBlocks'][0]['urls'][0]['result']['args'][0]['value']
		body=''
		#To be refined when more than one target is defined
			#body='\x0a'+body+url['result']['args'][0]['value'] 
		body="Tap targets are too close to others: "+'\x0a'+target
		ws.write(index,7,body,style)
		
	#UseLegibleFontSizes
	if (jsonO['formattedResults']['ruleResults']['UseLegibleFontSizes']['ruleImpact'] == 0):
		ws.write(index,8,'N/A',greenStyle)
	else:
		if(jsonO['formattedResults']['ruleResults']['UseLegibleFontSizes']['ruleImpact'] < impactRedThreshold):
			style=orangeStyle
		else:
			style=redStyle
		body='Avoid small font sizes, too hard to read: '+'\x0a'+'Text strings too small: '
		for url in jsonO['formattedResults']['ruleResults']['UseLegibleFontSizes']['urlBlocks'][0]['urls']:
			name=url['result']['args'][0]['value']
			body=body+'\x0a'+name
		ws.write(index,8,body,style)

#get json payload from google pagespeed api
def getJson(apiCmd):
	req = urllib2.Request(apiCmd)
	try:
		resp = urlopen(req)
	except HTTPError as e:
		print 'The server couldn\'t fulfill the request.'
		print 'Error code: ', e.code
	except URLError as e:
		print 'We failed to reach a server.'
		print 'Reason: ', e.reason
	else:
		result=resp.read()
		return(result)

		
