#!/usr/bin/python
import sys
import zipfile
import datetime

# Define constants for the states of the iCal parser.
ICALSTART = 0
ICALINVEVENT = 1

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
	print("merge.py --week-to-view startDate noOfWeeks data.ics template.docx")
	sys.exit(0)
	
if sys.argv[1] == "--week-to-view":
	if len(sys.argv) == 6:
		startDate = datetime.datetime.strptime(sys.argv[2], "%Y%m%d")
		if not startDate.weekday == 0:
			print("ERROR: Start date is not a Monday.")
			sys.exit(0)
		noOfWeeks = int(sys.argv[3])
		parseICalFile(sys.argv[4])
		with zipfile.ZipFile(sys.argv[5], "r") as templateDocx:
			textHandle = templateDocx.open("word/document.xml")
			docxText = str(textHandle.read())
			textHandle.close()
			weekToView = docxText[docxText.find("<w:body>")+8:docxText.find("</w:body>")]
			for calendarYear in sorted(calendar.keys()):
				print(calendarYear)
	else:
		print("ERROR: week-to-view - incorrect number of parameters.")
