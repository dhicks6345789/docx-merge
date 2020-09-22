#!/usr/bin/python

# DOCX Merge - a utility to do document merges with DOCX files.

# Standard libraries.
import os
import re
import sys
import shutil
import zipfile
import datetime

theTimezone = datetime.timezone("Europe/London")
#london_dt = aware_dt.astimezone(pytz.timezone('Europe/London'))

# The python-docx library, for manipulating DOCX files.
# Importantly, when installing with pip, that not the "docx" library, that an earlier version - do "pip install python-docx".
import docx

# Possible states for the iCal parser.
ICALSTART = 0
ICALINVEVENT = 1

# Used for placeholder names in calendars.
DAYNAMES = ["MO","TU","WE","TH","FR","SA","SU"]
DAYTITLES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

# DOCX files are ZIP files - we need a folder to unzip the contenst into if we want to modify a contained file.
TEMPLATETEMP = "templateTemp/"

ONEDAY = datetime.timedelta(days=1)

# A place to put calendar data read from an iCal file.
calendar = {}

# Make sure the given Year exists in the calendar.
def addCalendarYear(theYear):
	if not theYear in calendar.keys():
		calendar[theYear] = {}

# Make sure the given Month exists in the calendar.
def addCalendarMonth(theYear, theMonth):
	addCalendarYear(theYear)
	if not theMonth in calendar[theYear].keys():
		calendar[theYear][theMonth] = {}

# Make sure the given Day exists in the calendar.		
def addCalendarDay(theYear, theMonth, theDay):
	addCalendarMonth(theYear, theMonth)
	if not theDay in calendar[theYear][theMonth].keys():
		calendar[theYear][theMonth][theDay] = []

# Add the given Item to the calendar, at the given day / month / year.
def addCalendarItem(theYear, theMonth, theDay, theItem):
	addCalendarDay(theYear, theMonth, theDay)
	calendar[theYear][theMonth][theDay].append(theItem)

# Strip unwanted characters from strings.
def normaliseString(theString):
	result = ""
	for resultItem in theString.replace("\\n","\n").replace("\\,",",").replace("·","").split("\n"):
		result = result + resultItem.strip() + "\n"
	return(result.strip())

def unZeroPad(theString):
	if theString[0] == "0":
		return(theString[1:])
	return(theString)

def time24To12Hour(theString):
	result = theString
	matchResult = re.match("(\d\d):(\d\d): ", theString)
	if not matchResult == None:
		hour = int(matchResult.group(1))
		minuteString = ""
		if not matchResult.group(2) == "00":
			minuteString = ":" + str(int(matchResult.group(2)))
		if hour == 12 and matchResult.group(0) == "00":
			result = "Midday: "
		elif hour > 12:
			result = str(hour-12) + minuteString + "pm: "
		else:
			result = str(hour) + minuteString + "am: "
		result = result + theString[7:]
	return(result)

# A basic iCal parser.
def parseICalFile(theFilename):
	iCalState = ICALSTART
	iCalData = {}
	
	# Read the iCal file in as a bunch of text entries. We don't just use readlines() as some entries can be split over multiple lines, so we have to detect those
	# and stick them back together as we go along.
	iCalLines = []
	iCalHandle = open(theFilename, encoding="utf-8")
	for iCalLine in iCalHandle:
		if iCalLine.startswith(" "):
			iCalLines[-1] = iCalLines[-1] + iCalLine[1:].rstrip()
		else:
			iCalLines.append(iCalLine.rstrip())
	iCalHandle.close()
	
	# Now, parse the lines read above into calendar data.
	iCalBlock = ""
	for iCalLine in iCalLines:
		if iCalState == ICALSTART and iCalLine.startswith("BEGIN:VEVENT"):
			iCalState = ICALINVEVENT
			iCalData = {}
			iCalBlock = ""
		if iCalState == ICALINVEVENT:
			iCalBlock = iCalBlock + iCalLine + "\n"
		if iCalState == ICALINVEVENT and iCalLine.startswith("DTSTART:"):
			iCalData["StartDate"] = datetime.datetime.strptime(iCalLine.split(":",1)[1].split("T")[0], "%Y%m%d")
			iCalData["StartTime"] = datetime.datetime.strptime(iCalLine.split(":",1)[1].split("T")[1].split("Z")[0], "%H%M%S")
		if iCalState == ICALINVEVENT and iCalLine.startswith("DTSTART;VALUE=DATE:"):
			iCalData["StartDate"] = datetime.datetime.strptime(iCalLine.split(":",1)[1], "%Y%m%d")
		if iCalState == ICALINVEVENT and iCalLine.startswith("DTEND:"):
			iCalData["EndDate"] = datetime.datetime.strptime(iCalLine.split(":",1)[1].split("T")[0], "%Y%m%d")
			iCalData["EndTime"] = datetime.datetime.strptime(iCalLine.split(":",1)[1].split("T")[1].split("Z")[0], "%H%M%S")
		if iCalState == ICALINVEVENT and iCalLine.startswith("DTEND;VALUE=DATE:"):
			iCalData["EndDate"] = datetime.datetime.strptime(iCalLine.split(":",1)[1], "%Y%m%d")
		if iCalState == ICALINVEVENT and iCalLine.startswith("DESCRIPTION:"):
			iCalData["Description"] = iCalLine.split(":",1)[1]
		if iCalState == ICALINVEVENT and iCalLine.startswith("END:VEVENT"):
			iCalState = ICALSTART
			if "StartDate" in iCalData.keys():
				# Do we not have an EndDate set? Assume the event lasts one day.
				if not "EndDate" in iCalData.keys():
					iCalData["EndDate"] = iCalData["StartDate"]
				# Does the event not have a start / end time set but seems to last from one day until the next? Simply the event into lasting one day.
				if iCalData["EndDate"] == (iCalData["StartDate"] + ONEDAY):
					if not "StartTime" in iCalData.keys() and not "EndTime" in iCalData.keys():
						iCalData["EndDate"] = iCalData["StartDate"]
				timeString = ""
				if "StartTime" in iCalData.keys():
					timeString = iCalData["StartTime"].strftime("%H:%M: ")
				eventLength = iCalData["EndDate"] - iCalData["StartDate"]
				currentDate = iCalData["StartDate"]
				for eventDay in range(0, eventLength.days+1):
					addCalendarItem(currentDate.year, currentDate.month, currentDate.day, timeString + normaliseString(iCalData["Description"]))
					currentDate = currentDate + ONEDAY
			else:
				print("Unhandled event:\n" + iCalBlock.strip())

# Extract the given DOCX file to a given temporary folder.
# Reads and returns the contents of the "document.xml" contained in the DOCX file.
def extractDocx(theFilename, destinationPath):
	templateDocx = zipfile.ZipFile(theFilename, "r")
	templateDocx.extractall(destinationPath)
	templateDocx.close()
	textHandle = open(destinationPath + "word/document.xml")
	docxText = str(textHandle.read())
	textHandle.close()
	return(docxText)

# Turns the contents of the given folder into a DOCX file. Deletes the source folder when done.
def compressDocx(sourcePath, theFilename):
	theDocx = zipfile.ZipFile(sys.argv[6], "w")
	for root, dirs, files in os.walk(sourcePath):
		for file in files:
			theDocx.write(os.path.join(root, file), os.path.join(root, file)[len(sourcePath):])
	theDocx.close()
	shutil.rmtree(sourcePath)

# Writes a file to the given path.
def putFile(thePath, theData):
	textHandle = open(thePath, "w")
	textHandle.write(theData)
	textHandle.close()
											      
# Check arguments, print a usage message if needed.
if len(sys.argv) == 1:
	print("DOCX Merge - merges data into DOCX templates. Usage:")
	print("merge.py --week-to-view startDate noOfWeeks data.ics template.docx output.docx")
	sys.exit(0)

# The user wants a week-to-view calendar.
if sys.argv[1] == "--week-to-view":
	if len(sys.argv) == 7:
		# Check the start date is a Monday.
		startDate = datetime.datetime.strptime(sys.argv[2], "%Y%m%d")
		if not startDate.weekday() == 0:
			print("ERROR: Start date is not a Monday.")
			sys.exit(0)
		# Figure out the number of weeks to produce a calendar for.
		noOfWeeks = int(sys.argv[3])
		# Read the calendar data.
		parseICalFile(sys.argv[4])
		
		# The python-docx library doesn't have a function to duplicate pages, so we do that part ourselves by duplicating the main body of XML from
		# the "document.xml" contained in the DOCX file.
		docxText = extractDocx(sys.argv[5], TEMPLATETEMP)
		bodyStart = docxText.find("<w:body>")+8
		bodyEnd = docxText.find("</w:body>")
		newDocxText = docxText[:bodyStart]
		# Copy the main body text once for each week the user wants to show. Might be multiple pages.
		for week in range(0, noOfWeeks):
			weekToViewText = docxText[bodyStart:bodyEnd]
			for weekDay in range(0, 7):
				today = startDate + datetime.timedelta(days=(week*7)+weekDay)
				# Find the "title" string for the day.
				weekToViewText = weekToViewText.replace("{{" + DAYNAMES[weekDay] + "TI}}", today.strftime("%A, ") + unZeroPad(today.strftime("%d")) + today.strftime(" %B"))
				# Find the "content" string for the day.
				weekToViewText = weekToViewText.replace("{{" + DAYNAMES[weekDay] + "CO}}", "{{" + DAYNAMES[weekDay] + "-WEEK" + str(week) + "}}")
			newDocxText = newDocxText + weekToViewText
		# Re-write the content back to the output location.
		newDocxText = newDocxText + docxText[bodyEnd:]
		putFile(TEMPLATETEMP + "word/document.xml", newDocxText)
		compressDocx(TEMPLATETEMP, sys.argv[6])
		
		# Now, read the output file again with the python-docx library.
		templateDocx = docx.Document(sys.argv[6])
		for week in range(0, noOfWeeks):
			for weekDay in range(0, 7):
				dayContents = ""
				today = startDate + datetime.timedelta(days=(week*7)+weekDay)
				if today.year in calendar.keys():
					if today.month in calendar[today.year].keys():
						if today.day in calendar[today.year][today.month].keys():
							for dayItem in sorted(calendar[today.year][today.month][today.day]):
								dayContents = dayContents + time24To12Hour(dayItem.replace("\n",", ")) + "\n"
				dayContents = dayContents.strip()
				dayString = "{{" + DAYNAMES[weekDay] + "-WEEK" + str(week) + "}}"
				for paragraph in templateDocx.paragraphs:
					if dayString in paragraph.text:
						paragraph.text = dayContents
				for table in templateDocx.tables:
					for row in table.rows:
						for cell in row.cells:
							for paragraph in cell.paragraphs:
								if dayString in paragraph.text:
									paragraph.text = dayContents
		# Write out the final version of the DOCX file.
		templateDocx.save(sys.argv[6])
	else:
		print("ERROR: week-to-view - incorrect number of parameters.")
