#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Spreadsheet.py
# Version:  1.2.1
# Author:   Glenn Abastillas
# Date:     August 21, 2015
#
# Purpose: Allows the user to:
#           1.) Load a spreadsheet into memory.
#           2.) Transpose columns and rows.
#           3.) Find a search term and return the column and row it is located in.
#
# This class does not have scripting code in place.
#
# This class is directly inherited by the following classes:
#       - SpreadsheetPlus.py
#
# Updates:
# 1. [2015/12/04] - added: method open().
# 2. [2016/02/09] - in load() method, removed lines 125 - 128 including else-statement. Moved file_in from inside nested else-statement. Version changed to 1.2.1.
# 3. [2016/02/29] - changed wording of notes in line 17 from '... class is used in the following ...' to '... class is directly inherited by the following ...'.
# 4. [2016/03/07] - removed unused method - def determineGroupByRange(self, columnToBeDetermined = 0, rangeToBeDetermined = 0)
# - - - - - - - - - - - - -
"""	Creates a manipulable spreadsheet object from a text file.

The Spreadsheet class is used to represent text files as objects for further
use by other classes. This is a base class and does not inherit from other
classes. Two methods exist in this class that can be used as static methods:
(1) open(doc): open a specified document and (2) save().

"""
__author__      = "Glenn Abastillas"
__copyright__   = "Copyright (c) August 21, 2015"
__credits__     = "Glenn Abastillas"

__license__     = "Free"
__version__     = "1.2.1"
__maintainer__  = "Glenn Abastillas"
__email__       = "a5rjqzz@mmm.com"
__status__      = "Deployed"

class Spreadsheet(object):

	def __init__(self, filePath = None, savePath = None):
		"""	Initialize an instance of this class
			@param  filePath: path of the spreadsheet file to be loaded
			@param	savePath: write to this location
		"""

		self.spreadsheet = list()		#list containing spreadsheet

		self.filePath	 = filePath		#location of the spreadsheet
		self.savePath	 = savePath		#location of the spreadsheet

		self.loaded		 = False		#spreadsheet loaded?
		self.initialized = False		#spreadsheet initialized?
		self.transposed	 = False		#checks if self.spreadsheet stores rows (=False) or columns (=True)

		self.iter_index	 = 0			# Used for next() function

		if filePath is not None:
			self.initialize(filePath)

	def __getitem__(self, key):
		"""	Enable list[n] syntax
			@param  key: index number of item.
		"""
		return self.spreadsheet[key]
	
	def __iter__(self):
		"""	Enable iteration of this class
		"""
		return self

	def __len__(self):
		"""	Return number of items in spreadsheet
		"""
		return len(self.spreadsheet)

	def next(self):
		"""	Get next item in iteration
		"""
		try:
			self.iter_index += 1
			return self.spreadsheet[self.iter_index - 1]
		except(IndexError):
			self.iter_index = 0
			raise StopIteration
	
	def initialize(self, filePath = None, sep = "\t"):
		"""	Open the file and parse out rows and columns
			@param	filePath: spreadsheet file to load into memory
		"""
		
		openedFilePath	 = self.open(filePath).splitlines()

		for line in openedFilePath:
			self.spreadsheet.append(line.split(sep))

		self.filePath 	 = filePath
		self.loaded 	 = True

	def open(self, filePath):
		"""	Opens an indicated text file for processing
			@param	filePath: path of file to load
			@return	String of opened text file
		"""
		fileIn1 = open(filePath, 'r')
		fileIn2 = fileIn1.read()
		fileIn1.close()
		return fileIn2

	def save(self, savePath=None, saveContent=None, saveType='w'):
		"""	Write content out to a file
			@param	savePath: name of the file to be saved
			@param	saveContent: list of rows/columns to be saved
			@param	saveType: indicate overwrite ('w') or append ('a')
		"""
		if saveContent is None:
			saveContent = self.prepareForSave(self.spreadsheet)
		else:
			saveContent = self.prepareForSave(saveContent)

		saveFile = open(savePath, saveType)
		saveFile.write(saveContent)
		saveFile.close()

	def prepareForSave(self, spreadsheet=None):
		"""	Prepare the spreadsheet for saving
			@param	spreadsheet: list of rows/columns to prepare
			@return	String of spreadsheet in normal form (e.g., not transposed)
		"""
		rowList = list()
		append	= rowList.append

		for row in spreadsheet:
			append("\t".join(row))

		saveContent = "\n".join(rowList)

		return saveContent
	
	def getSavePath(self):
		"""	Get the save path
			@return String of the save path
		"""
		return self.savePath

	def getFilePath(self):
		"""	Get the file path
			@return	String of the file path
		"""
		return self.filePath

	def getSpreadsheet(self):
		"""	Get this Spreadsheet
			@return	List of rows and columns with content
		"""
		return self.spreadsheet

	def setSavePath(self, savePath):
		"""	Set the location for saved files
			@param	savePath: location to store saved files
		"""
		self.savePath = savePath

	def setFilePath(self, filePath):
		"""	Set this object to a new file
			@param	filePath: location to new file
		"""
		self.initialize(filePath)

	def toString(self, fileToString = None):
		"""	Print input to screen
			@param fileToString: file to print out as string to screen.
		"""
		
		if fileToString is None:
			fileToString = self.spreadsheet
			
		if self.initialized:
			join = str.join

			# Loop through the lines to join them as strings
			for line in fileToString:
				print join("\t\t", line)
		else:
			for line in fileToString:
				print line

		print "\n\n"
		
	def transpose(self):
		"""	Transpose this spreadsheet's rows and columns
		"""
		temporarySpreadsheet = list()
					
		# CALL THESE JUST ONCE BEFORE LOOP(S)
		append		= temporarySpreadsheet.append
		longestItem = len(max(self.spreadsheet, key = len))

		# Loop through the longest row (if transposed=False) or column (if transposed = True)
		for index in xrange(longestItem):

			# At this index, insert a list for the new row/column
			append(list())

			# CALL THESE JUST ONCE BEFORE LOOP(S)
			append2 = temporarySpreadsheet[index].append

			# Loop through current spreadsheet to transpose rows<==>columns
			for line in self.spreadsheet:
				try:
					append2(line[index])
				except(IndexError):

					# If the specified cell does not exist, i.e., blank
					append2("")

		self.spreadsheet = temporarySpreadsheet
		self.transposed = not(self.transposed)

	def reset(self):
		"""	Reset all data in this class
		"""

		self.spreadsheet = list()	#list containing spreadsheet

		self.filePath	 = None		#location of the spreadsheet
		self.savePath	 = None		#location of the spreadsheet

		self.loaded		 = False	#spreadsheet loaded?
		self.initialized = False	#spreadsheet initialized?
		self.transposed	 = False	#checks if self.spreadsheet stores rows (=False) or columns (=True)

		self.iter_index	 = 0		# Used for next() function

if __name__=="__main__":

	d = Spreadsheet("../files/debridementSamples.txt")
	print d.getSpreadsheet()
	d.transpose()
	print d.getSpreadsheet()