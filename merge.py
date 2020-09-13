#!/usr/bin/python
import os
import sys
import shutil
import zipfile
import datetime

# Define constants for the states of the iCal parser.
ICALSTART = 0
ICALINVEVENT = 1

DAYNAMES = {0:"MO",1:"TU",2:"WE",3:"TH",4:"FR",5:"SA",6:"SU"}

DATETIMEFORMAT = "%Y%m%dT%H%M%SZ"

calendar = {}

def addCalendarYear(theYear):
	if not theYear in calendar.keys():
		calendar[theYear] = {}

def addCalendarMonth(theYear, theMonth):
	addCalendarYear(theYear)
	if not theMonth in calendar[theYear].keys():
		calendar[theYear][theMonth] = {}
		
def addCalendarDay(theYear, theMonth, theDay):
	addCalendarMonth(theYear, theMonth)
	if not theDay in calendar[theYear][theMonth].keys():
		calendar[theYear][theMonth][theDay] = []
		
def addCalendarItem(theYear, theMonth, theDay, theItem):
	addCalendarDay(theYear, theMonth, theDay)
	calendar[theYear][theMonth][theDay].append(theItem)

def parseICalFile(theFilename):
	iCalState = ICALSTART
	iCalData = {}
	iCalHandle = open(theFilename)
	for iCalLine in iCalHandle.readlines():
		iCalLine = iCalLine.strip()
		if iCalState == ICALSTART and iCalLine.startswith("BEGIN:VEVENT"):
			iCalState = ICALINVEVENT
			iCalData = {}
		elif iCalState == ICALINVEVENT and iCalLine.startswith("DTSTART:"):
			iCalData["StartDate"] = iCalLine.split(":")[1]
		elif iCalState == ICALINVEVENT and iCalLine.startswith("DTEND:"):
			iCalData["EndDate"] = iCalLine.split(":")[1]
		elif iCalState == ICALINVEVENT and iCalLine.startswith("DESCRIPTION:"):
			iCalData["Description"] = iCalLine.split(":")[1]
		elif iCalState == ICALINVEVENT and iCalLine.startswith("END:VEVENT"):
			iCalState = ICALSTART
			if "StartDate" in iCalData.keys() and "EndDate" in iCalData.keys():
				startDate = datetime.datetime.strptime(iCalData["StartDate"], DATETIMEFORMAT)
				endDate = datetime.datetime.strptime(iCalData["EndDate"], DATETIMEFORMAT)
				eventLength = endDate - startDate
				if eventLength.days == 0:
					addCalendarItem(startDate.year, startDate.month, startDate.day, iCalData["Description"])
	iCalHandle.close()

if len(sys.argv) == 1:
	print("DOCX Merge - merges data into DOCX templates. Usage:")
	print("merge.py --week-to-view startDate noOfWeeks data.ics template.docx output.docx")
	sys.exit(0)
	
if sys.argv[1] == "--week-to-view":
	if len(sys.argv) == 7:
		startDate = datetime.datetime.strptime(sys.argv[2], "%Y%m%d")
		if not startDate.weekday() == 0:
			print("ERROR: Start date is not a Monday.")
			sys.exit(0)
		noOfWeeks = int(sys.argv[3])
		parseICalFile(sys.argv[4])
		
		templateDocx = zipfile.ZipFile(sys.argv[5], "r")
		templateDocx.extractall("templateTemp")
		templateDocx.close()
		textHandle = open("templateTemp/word/document.xml")
		docxText = str(textHandle.read())
		textHandle.close()
		bodyStart = docxText.find("<w:body>")+8
		bodyEnd = docxText.find("</w:body>")
		newDocxText = docxText[:bodyStart]
		for week in range(0, noOfWeeks):
			weekToViewText = docxText[bodyStart:bodyEnd]
			for weekDay in range(0, 5):
				dayString = "{{" + DAYNAMES[weekDay] + "1}}"
				today = startDate + datetime.timedelta(days=(week*7)+weekDay)
				if today.year in calendar.keys():
					if today.month in calendar[today.year].keys():
						if today.day in calendar[today.year][today.month].keys():
							weekToViewText.replace(dayString, str(calendar[today.year][today.month][today.day]))
			newDocxText = newDocxText + weekToViewText
		newDocxText = newDocxText + docxText[bodyEnd:]
		textHandle = open("templateTemp/word/document.xml", "w")
		textHandle.write(newDocxText)
		textHandle.close()
		
		templateDocx = zipfile.ZipFile(sys.argv[6], "w")
		for root, dirs, files in os.walk("templateTemp/*"):
			for file in files:
				templateDocx.write(os.path.join(root, file))
		templateDocx.close()
	else:
		print("ERROR: week-to-view - incorrect number of parameters.")
