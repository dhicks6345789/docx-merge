#!/usr/bin/python
import sys
import datetime

# Define constants for the states of the iCal parser.
ICALSTART = 0
ICALINVEVENT = 1

DATETIMEFORMAT = "%Y%m%dT%H%M%SZ"

calendar = {}

def addCalendarYear(theYear):
	if not theYear in calendar.keys():
		calendar[theYear] = {}

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
					addCalendarYear(startDate.year)
					#print(iCalData)
					#print("2020 11 27 T150000Z")
					# Code goes here - add the event to the data.
	iCalHandle.close()

if len(sys.argv) == 1:
	print("DOCX Merge - merges data into DOCX templates. Usage:")
	print("merge.py --week-to-view data.ical template.docx")
	sys.exit(0)
	
if sys.argv[1] == "--week-to-view":
	if len(sys.argv) == 4:
		parseICalFile(sys.argv[2])
		print(calendar)
		#print(sys.argv[3])
	else:
		print("ERROR: week-to-view - incorrect number of parameters.")
