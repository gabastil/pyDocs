#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name: 	SpreadsheetSearchLog.py
# Version: 	1.1.0
# Author: 	Glenn Abastillas
# Date: 	October 20, 2015
#
# Purpose: Allows the user to:
#           1.) Log times for each run of SpreadsheetSearch
#
# To see the script run, go to the bottom of the page.
#
# This class is directly inherited by the following classes:
#	- Analyze.py
#
# Updates:
# 1. [2016/02/29] - added comments/documentation to the methods.
# 2. [2016/02/29] - changed wording of notes in line 14 from '... class is used in the following ...' to '... class is directly inherited by the following ...'.
#
# - - - - - - - - - - - - -
"""records the run time(s) for client sites that underwent CAPD-/DICE-Tuning.
"""
__author__      = "Glenn Abastillas"
__copyright__   = "Copyright (c) October 20, 2015"
__credits__     = "Glenn Abastillas"

__license__     = "Free"
__version__     = "1.1.0"
__maintainer__  = "Glenn Abastillas"
__email__       = "a5rjqzz@mmm.com"
__status__      = "Development"

import os, time

class SpreadsheetSearchLog(object):

	def openLog(self, filePath = "files\\", docName = "SpreadsheetSearchLog.txt"):
		"""	open the SpreadsheetSearch Log file for editing.

			filePath --> base file path containing the log
			docName  --> name of the SpreadsheetSearch log
		"""

		logFile = open(filePath + docName, 'r')
		log 	= logFile.read().split('\n')
		logFile.close()

		return log

	def saveLog(self, log = "", filePath = "files\\", docName = "SpreadsheetSearchLog.txt"):
		"""	save the SpreadsheetSearch log data to the specified file.

			log 	 --> data to be written to the file
			filePath --> base file path containing the log
			docName  --> name of the SpreadsheetSearch log
		"""

		logFile = open(filePath + docName, 'w')
		logFile.write(log)

		logFile.close()

	def updateLog(self, art = 0.00, trt = 0.00, dcr = "None", pfn = "None", filePath = "files\\", docName = "SpreadsheetSearchLog.txt"):
		"""	update the SpreadsheetSearch log information after extracting and 
			searching the documents.

			art 	 --> average run time
			trt 	 --> total run time
			dcr 	 --> time to complete one DICE Code
			pfn 	 --> name of processed document
			filePath --> base file path containing the log
			docName  --> name of the SpreadsheetSearch log
		"""

		timeStamp = time.strftime("%Y/%m/%d\t[%H:%M:%S]")
		avgRunTim = art
		totRunTim = trt
		diceCodes = dcr
		procFileN = pfn

		newLine = '\t'.join([timeStamp, str(avgRunTim) + "\"", str(totRunTim) + "\"", procFileN, str(diceCodes)])

		log = self.openLog(filePath = filePath, docName = docName)
		log = [newLine] + log
		log = '\n'.join(log)

		self.saveLog(log = log, filePath = filePath, docName = docName)

if __name__ == "__main__":

	ssl = SpreadsheetSearchLog()
	ssl.updateLog()
