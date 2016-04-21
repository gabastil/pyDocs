#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Spreadsheet.py
# Version:  1.2.2
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
# 2. [2016/02/09] - in load() method, removed lines 125 - 128 including else-statement. Moved file_in from inside nested else-statement. Version changed to 1.2.2.
# 3. [2016/02/29] - changed wording of notes in line 17 from '... class is used in the following ...' to '... class is directly inherited by the following ...'.
# 4. [2016/03/07] - removed unused method - def determineGroupByRange(self, columnToBeDetermined = 0, rangeToBeDetermined = 0)
# 5. [2016/04/18] - added: addColumn(), newColumn() and fillColumn() methods
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
__version__     = "1.2.2"
__maintainer__  = "Glenn Abastillas"
__email__       = "a5rjqzz@mmm.com"
__status__      = "Deployed"

class Spreadsheet(object):

	def __init__(self, filePath=None, savePath=None, columns=["columnName"]):
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
		else:
			self.spreadsheet.extend([columns])

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
	
	def addColumn(self, newColumn=None):
		"""	Add a column to the spreadsheet
			@param	newColumn: (list) column to add
		"""
		# If newColumn is none, raise an error
		if newColumn is None:
			raise ValueError("Please specify a new column list to add to the spreadsheet. E.g., Spreadsheet.addColumn(columnAsList)")

		# Transpose spreadsheet to edit column
		self.toColumns()
		self.spreadsheet.append(newColumn)
		self.refresh()

	def addToColumn(self, index=-1, content=None):
		"""	Fill a column with text
			@param	content: content to insert
			@param	index: column to fill
		"""
		self.toColumns()
		self.spreadsheet[index].append(content)

	def newColumn(self, name=" ", fillWith=" "):
		"""	Add a new (empty) column to the spreadsheet
			@param	name: name of column
			@param	fillWith: contents of new column
		"""
		# Transpose spreadsheet to edit column
		self.toColumns()

		newColumn 	 = [fillWith] * len(self.spreadsheet[0])
		newColumn[0] = name

		self.spreadsheet.append(newColumn)
		self.refresh()

	def fillColumn(self, column=-1, fillWith=" ", skipTitle=True):
		"""	Fill a column with text
			@param	column: column to fill (index or column name)
			@param	fillWith: contents of new column
		"""
		# Transpose spreadsheet to edit column
		self.toColumns()

		initialIndex = 0

		newColumn	 = list()
		append		 = newColumn.append

		if type(column) == type(str()):
			self.toRows()
			column = self.spreadsheet[0].index(column)
			self.toColumns()

		if skipTitle==True:
			append(self.spreadsheet[column][0])
			initialIndex = 1

		for cell in self.spreadsheet[column][initialIndex:]:
			append(fillWith)

		self.spreadsheet[column] = newColumn

	def rename(self, name=None, index=-1):
		"""	Rename a specified column
			@param	name: name for the column
			@param	index: column index
		"""
		# Transpose spreadsheet to edit column
		self.toColumns()

		self.spreadsheet[index][0] = name

	def setCell(self, row, column, content):
		""" Set a cell to content
			@param	content: content to fill cell
		"""
		self.toRows()
		self.spreadsheet[row][column]=content

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
			saveContent = self.prepareForSave()
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
		# Use Spreadsheet's spreadsheet if none specified
		if spreadsheet is None:
			self.toRows()
			spreadsheet = self.spreadsheet

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

	def toColumns(self):
		"""	Transpose to columns
		"""
		# Transpose spreadsheet to edit column
		if not self.transposed:
			self.transpose()

	def toRows(self):
		"""	Transpose to rows
		"""
		# Transpose spreadsheet to edit rows
		if self.transposed:
			self.transpose()

	def sort(self, column=0, reverse=False, hasTitle=True):
		""" Sort spreadsheet based on column
			@param	column: column to sort by
			@param	reverse: reverse sort
		"""
		self.toRows()
		if hasTitle == True:
			header = self.spreadsheet[0]
			body = sorted(self.spreadsheet[1:], key=lambda row:row[column], reverse=reverse)
			self.spreadsheet = [header] + body
		else:
			self.spreadsheet = sorted(self.spreadsheet, key=lambda row:row[column], reverse=reverse)

	def refresh(self):
		""" Make sure all columns/rows are the same length
		"""
		self.transpose()
		self.transpose()

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

	#d = Spreadsheet("../files/debridementSamples.txt")
	d = Spreadsheet("../files/test.txt")
	print d.getSpreadsheet()
	d.transpose()
	print d.getSpreadsheet()
	d.fillColumn(0,"filled")
	print d.spreadsheet
	d.addColumn([1,2,3])
	print d.spreadsheet
	d.fillColumn(-1,"filled")
	print d.spreadsheet
	d.addColumn([1,2,3,4,5])
	print d.spreadsheet
	d.addToColumn(2,"test")
	print d.spreadsheet
	print d.spreadsheet

	p = Spreadsheet(columns=["blah", "booh", "bah", "ble"] + ["too", "me"])
	p.newColumn("blech")
	p.newColumn()
	print "P:\t", p.spreadsheet
	p.transpose()
	print "P:\t", p.spreadsheet
	p.transpose()
	print "P:\t", p.spreadsheet

	print p.prepareForSave()
	print p.addToColumn(0,87)
	print p.addToColumn(0,55)
	print p.addToColumn(0,785)
	print p.addToColumn(1,3)
	p.refresh()
	print "FIRSTP:\t", p.spreadsheet
	p.setCell(3,2,"SETSELL")
	p.fillColumn("too", "can")
	p.addToColumn(4,"acd")
	print "P:\t", p.spreadsheet
	p.fillColumn("ble", "filledWithThis")
	print "P:\t", p.spreadsheet
	p.toRows()
	print "ddP:\t", p.spreadsheet
	print p.sort(4)
	print "P:\t", p.spreadsheet
	print p.sort(4, hasTitle=False)
	print "P:\t", p.spreadsheet
