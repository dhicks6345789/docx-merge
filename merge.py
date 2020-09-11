#!/usr/bin/python
import os
import io
import sys
import json
import time
import shutil
import pandas
import dataLib
import datetime

# PyPDFs - used for merging existing PDF documents.
import PyPDF2



# Load the config file (set by the system administrator).
config = dataLib.loadConfig(["dataFolder"])

# Make sure the input / output folder exists.
calendarsFolder = config["dataFolder"] + os.sep + "Calendars"
os.makedirs(calendarsFolder, exist_ok=True)

print("rclone sync to local of one Word (Docs) doc")
print("rclone sync to local of two PDF documents")
print("HTTP get of iCal file")
print("MergeCalendar")
print("rclone upload of one DOCX file")
print("rclone download of one PDF document")
print("join of 3 PDF documents")
print("rclone upload of finished PDF document")

pdfsToMerge = []

# Check to see if there is content to merge.
frontMatterPath = calendarsFolder + os.sep + outputFilename + os.sep + "frontMatter.pdf"
if os.path.exists(frontMatterPath):
	print("Found front matter...")
	pdfsToMerge.append(frontMatterPath)

pdfsToMerge.append("temp.pdf")

backMatterPath = calendarsFolder + os.sep + outputFilename + os.sep + "backMatter.pdf"
if os.path.exists(backMatterPath):
	print("Found back matter...")
	pdfsToMerge.append(backMatterPath)
	
pdfMerger = PyPDF2.PdfFileMerger()
for pdfToMerge in pdfsToMerge:
	pdfMerger.append(pdfToMerge)
pdfMerger.write(calendarsFolder + os.sep + outputFilename + ".pdf")
