#!/usr/bin/python
import sys

if len(sys.argv) == 1:
	print("DOCX Merge - merges data into DOCX templates. Usage:")
	print("merge.py --week-to-view data.ical template.docx")
	sys.exit(0)
	
if sys.argv[1] == "--week-to-view":
	if len(sys.argv) == 4:
		print(sys.argv[2])
		print(sys.argv[3])
	else:
		print("ERROR: week-to-view - incorrect number of parameters.")
