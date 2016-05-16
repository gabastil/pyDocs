#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     SpreadsheetPlus.py
# Version:  1.0.0
# Author:   Glenn Abastillas
# Date:     September 21, 2015
#
# Purpose: Allows the user to:
#           1.) Load a spreadsheet into memory.
#           2.) Transpose columns and rows.
#           3.) Find a search term and return the column and row it is located in.
#
# To see the script run, go to the bottom of this page.
#
# This class is used in the following classes:
#   - SpreadsheetSearch.py
#   - SpreadsheetCompare.py
# 
# Updates:
# 1. [2015/12/03] added "savePath" variable to save() function.
# 2. [2015/12/04] optimized processes for speed, added saveFile() method.
# 3. [2015/12/07] optimized name creation in save() method.
# - - - - - - - - - - - - -
""" create a Spreadsheet object for two spreadsheet inputs that enables the user to manipulate both

The inherited Spreadsheet object allows for the creation of ---
"""
__author__      = "Glenn Abastillas"
__copyright__   = "Copyright (c) August 21, 2015"
__credits__     = "Glenn Abastillas"

__license__     = "Free"
__version__     = "1.0.0"
__maintainer__  = "Glenn Abastillas"
__email__       = "a5rjqzz@mmm.com"
__status__      = "Deployed"

from Spreadsheet import Spreadsheet

class SpreadsheetPlus(Spreadsheet):

	def __init__(self, filePath1=None, filePath2=None, savePath=None):
		""" Initialize an instance of this class
			@param  filePath1: path of the first  spreadsheet file
			@param  filePath2: path of the second spreadsheet file
			@param  savePath: write to this location
		"""
		self.spreadsheetPlus= list()	#list containing spreadsheet (e.g., Drools)
		self.filePath2		= filePath2	#location of the spreadsheet

		self.loadedPlus		= False		#spreadsheet loaded?
		self.initializedPlus= False		#spreadsheet initialized?
		self.transposedPlus	= False		#checks if self.spreadsheet stores rows (=False) or columns (=True)
		self.transformed    = False		#checks if self.spreadsheet was transformed
		
		self.oldSheet		= list()	#Stores old self.spreadsheet when this object is transformed
		self.oldSheetPlus	= list()	#Stores old self.spreadsheetPlus when this object is transformed

		super(SpreadsheetPlus, self).__init__(filePath=filePath1, savePath=savePath)
		print self.loaded

		if filePath2 is not None:
			self.initializePlus(filePath2)
	
	def initializePlus(self, filePath=None, sep="\t"):
		""" Open the file and parse out rows and columns
			@param	filePath: spreadsheet file to load into memory
		"""
		openedFilePath = self.open(filePath).splitlines()

		for line in openedFilePath:
			self.spreadsheetPlus.append(line.split(sep))

		self.filePath2 	= filePath
		self.loadedPlus	= True

	def toStringPlus(self, fileToString=None):
		""" Print out the input to screen
			@param	fileToString: string to print
		"""
		
		# If no input specified, use spreadsheetPlus
		if fileToString is None:
			fileToString = self.spreadsheetPlus
		
		# If
		if type(fileToString) == type(str()):
			return fileToString
		else:

			spreadsheetAsString = ""

			if self.initializedPlus:

				# CALL THESE JUST ONCE BEFORE LOOP(S)
				join    = str.join
				format  = str.format

				for line in fileToString:
					spreadsheetAsString = format("{0}{1}\n", spreadsheetAsString, join("\t", line))
			else:
				for line in fileToString:
					spreadsheetAsString = format("{0}{1}",   spreadsheetAsString, line)

			return spreadsheetAsString
		
	def transpose(self, sheet=1):
		""" Transpose a spreadsheet's rows and columns
			@param	sheet: spreadsheet to tranpose (1=spreadsheet, )
		"""
		# If 1, transpose the spreadsheet from filePath1
		if sheet == 1:
			super(SpreadsheetPlus, self).transpose()

		# If 2, transpose the spreadsheet from filePath2
		elif sheet == 2:
			temporary_spreadsheet = list()
			
			# CALL THESE JUST ONCE BEFORE LOOP(S)
			append = temporary_spreadsheet.append
			longest_row = len(max(self.spreadsheetPlus, key=len))

			# Loop through the longest row (if transposedPlus=False) or column (if transposedPlus=True)
			for index in xrange(longest_row):

				# At this index, insert a list for the new row/column
				append(list())

				# CALL THESE JUST ONCE BEFORE LOOP(S)
				append2 = temporary_spreadsheet[index].append

				# Loop through the lines of spreadsheetPlus to transpose rows and columns
				for line in self.spreadsheetPlus:
					try:
						append2(line[index])
					except(IndexError):
						append2("")

			self.spreadsheetPlus = temporary_spreadsheet
			self.transposedPlus	 = not(self.transposedPlus)

		# If 3, transpose both spreadsheets (i.e., filePath1, filePath2)
		elif sheet == 3:
			self.transpose(1)
			self.transpose(2)

		# If no number matches, return this error
		else:
			return ValueError("Please indicate spreadsheet 1 or 2. \
							   Enter 3 for both. (e.g., transpose(3))")

	def transform(self,*newColumns):
		"""	Create a new spreadsheet with specified columns
			@param	newColumns: columns to include from the old spreadsheet
		"""

		# Transpose spreadsheet to loop across columns
		if not self.transposed:
			self.transpose(1)

		# If no columns are specified, default to 1,4,5,6
		if len(newColumns) < 1:

			#Corresponds to columns: DICE-Code, pre-text, match, and post-text
			newColumns = (1,4,5,6)

		newSpreadsheet = list()
		append = newSpreadsheet.append

		# Loop through the specified columns
		for column in newColumns:
			append(self.spreadsheet[column])

		#--x self.oldSheet	 = self.spreadsheet
		self.spreadsheet = newSpreadsheet
		self.transformed = True

	def getSpreadsheetPlus(self):
		""" Get this Spreadsheet
			@return	List of rows and columns with content
		"""
		return self.spreadsheetPlus

	def save(self, sheet=0, savePath=None, saveContent=None, saveType='w'):
		"""	Write sheet to file
			@param	sheet: spreadsheet (1=spreadsheet;2=spreadsheetPlus)
			@param	savePath: name of the file to be saved
			@param	saveContent: list of rows/columns to be saved
			@param	saveType: indicate overwrite ('w') or append ('a')
		"""
		if sheet == 1:
			saveContent = self.prepareForSave(self.spreadsheet)
		elif sheet == 2:
			saveContent = self.prepareForSave(self.spreadsheetPlus)
		
		super(SpreadsheetPlus, self).save(savePath=savePath, saveContent=self.spreadsheet, saveType=saveType)

if __name__=="__main__":

	filePath1 = "../files/debridementSamples.txt"
	filePath2 = "../files/debridementSamples.txt"

	d = SpreadsheetPlus(filePath1, filePath2)
	print d.getSpreadsheet()
	d.transpose(1)
	print d.getSpreadsheet()