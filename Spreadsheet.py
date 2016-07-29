#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Spreadsheet.py
# Version:  1.0.0
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
# 2. [2016/02/09] - in load() method, removed lines 125 - 128 including else-statement. Moved file_in from inside nested else-statement.
# 3. [2016/02/29] - changed wording of notes in line 17 from '... class is used in the following ...' to '... class is directly inherited by the following ...'.
# 4. [2016/03/07] - removed unused method - def determineGroupByRange(self, columnToBeDetermined = 0, rangeToBeDetermined = 0)
# 5. [2016/04/18] - added: addColumn(), newColumn() and fillColumn() methods
# 5. [2016/04/21] - added: addColumn(), newColumn() and fillColumn() methods
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
__version__     = "1.0.0"
__maintainer__  = "Glenn Abastillas"
__email__       = "a5rjqzz@mmm.com"
__status__      = "Deployed"

class Spreadsheet(object):

	def __init__(self, filePath=None, savePath=None, delimiter="\t", columns=["columnName"]):
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
			self.initialize(filePath, delimiter)
		else:
			self.spreadsheet.extend([columns])

		self.COLUMN_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

	def __getitem__(self, key):
		"""	Enable list[n] syntax
			@param  key: index number of item.
		"""
		return self.spreadsheet[key]

	def __setitem__(self, i, item):
		"""	Enable list[n] = item assignment
			@param	i: index of item to replace
			@param	item: new item to assign
		"""
		self.spreadsheet[i] = item

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

	def addToColumn(self, column=-1, content=None):
		"""	Fill a column with text
			@param	column: column to fill (index or string)
			@param	content: content to insert
		"""
		index = self.getColumnIndex(column)

		self.toColumns()
		self.spreadsheet[index].append(content)

	def addRow(self, newRow=None):
		""" Add a row to the spreadsheet
			@param	newRow: (list) row to add
		"""
		# If newRow is none, raise an error
		if newRow is None:
			raise ValueError("Please specify a new row list to add to the spreadsheet. E.g., Spreadsheet.addRow(rowAsList)")

		# Transpose spreadsheet to edit row
		self.toRows()
		self.spreadsheet.append(newRow)
		self.refresh()

	def addToRow(self, row=-1, content=None):
		"""	append or extend data to a row
			@param	row:	 index of row to add data to
			@param	content: content to insert
		"""
		if type(content)==type(list()):
			self.spreadsheet[row].extend(content)
		else:
			self.spreadsheet[row].append(content)

	def fillColumn(self, column=-1, fillWith=" ", skipTitle=True, cellList=None):
		"""	Fill a column with text
			@param	column: column to fill (index or column name)
			@param	fillWith: contents of new column
			@param	cellList: list of tuples (cell, offset) to insert into the fillWith formula
			@param	skipTitle: start filling on the second row
		"""
		# Transpose spreadsheet to edit column
		self.toColumns()

		initialIndex = 0
		newColumn	 = list()
		append		 = newColumn.append

		# if column variable is a string (i.e., column name), get column with the same name's index	
		if type(column) == type(str()):
			column = self.getColumnIndex(column)

		# if skipTitle is True, start loop at index 1 not 0			
		if skipTitle==True:
			append(self.spreadsheet[column][0])
			initialIndex = 1

		# if there is a a cellList, there are values to replace
		if cellList != None:

			# items in cellList are tuples
			# the first part is the reference (e.g., "$A1" --> "$A{0}")
			# the second part is the row displacement, if any, (e.g., "-1")
			# for example, 	[("$A{0}", "1"), ("$C${0}", "-1")] becomes:
			# 				["$A2", "$C$0"] at i-index == 1 in the loop

			cellListLength = len(cellList)
			columnForLoop  = len(self.spreadsheet[column])

			# loop through the rows in the spreadsheet to fill content
			for i in range(columnForLoop)[initialIndex:]:

				fillWithToReplace = fillWith

				# loop through the list of cells to insert into the fillWith formula
				for j in xrange(cellListLength):

					cellListValue  = cellList[j][1]

					# if the cellList marker (i.e., {0}) to be replaced is a column
					# the reference value has to be a letter
					if type(cellListValue) == type(str()):
						referenceValue = self.COLUMN_ALPHA[column]

					# else, the reference value is a number
					else:
						referenceValue = str(i+cellListValue)

					# reference cell to insert into the fillWith content
					reference= cellList[j][0].replace("{0}", referenceValue)

					# marker indicates where the reference should go (e.g., "{0} goes here and then {1}")
					marker	 = "{}{}{}".format("{",str(j), "}")
					
					# replace the marker in the fillWith string with the reference cell
					# e.g., "{0} goes here and then {1}"; marker = "{0}"; reference = "$A$1"
					#		"$A$1 goes here and then {1}"
					fillWithToReplace = fillWithToReplace.replace(marker, reference)

				append(fillWithToReplace)

		# if there is no cellList, just fill all cells with the same content
		else:

			columnForLoop = self.spreadsheet[column][initialIndex:]

			# loop through the rows in the spreadsheet to fill content
			for row in columnForLoop:
				append(fillWith)

		self.spreadsheet[column] = newColumn

	def getColumn(self, column, number=False, header=False):
		"""	get column at specified index
			@param	column: index or name of desired column
			@param	number: return floats of numbers in column
			@param	header: include column header
			@return	list of column
		"""
		self.toColumns()

		if number==True:
			#print self.spreadsheet[self.getColumnIndex(column)][1:]
			column = [self.spreadsheet[0]] + [float(row) for row in self.spreadsheet[self.getColumnIndex(column)][1:] if row != '']
		else:
			column = self.spreadsheet[self.getColumnIndex(column)]

		if header==False:
			return column[1:]

		return column

	def getColumnIndex(self, column):
		"""	get the numerical index for a column
			@param	column:	column character (e.g., 'A') or column name
			@return	integer index of column specified
		"""


		# If 'column' is a number type, return the integer of that number
		if (type(column)==type(int())) or (type(column)==type(float())):
			return int(column)

		# If 'column' is a single character, return it's index
		elif len(column)==1:
			return self.COLUMN_ALPHA.index(column.upper())

		# If 'column' is a string, return it's index
		else:
			self.toRows()
			#print self.spreadsheet[0]
			#print self.spreadsheet[1]
			#print self.spreadsheet[2]
			columnIndex = self.spreadsheet[0].index(column)
			self.toColumns()
			return columnIndex

	def getColumnName(self, column):
		"""	get the name of specified column
			@param	column:	column character (e.g., 'A') or column name
			@return	string name of column
		"""
		index = self.getColumnIndex(column)
		return self.getColumn(column=column, header=True)[0]

	def getFilePath(self):
		"""	Get the file path
			@return	String of the file path
		"""
		return self.filePath

	def getHeaders(self):
		""" return headers for columns in the spreadsheet """
		self.toColumns()
		return [column[0] for column in self.spreadsheet]

	def getRow(self, index):
		"""	get row at specified index
			@param	index: index for desired row
			@return	list of row
		"""
		self.toRows()
		return self.spreadsheet[index+1]

	def getRowIndex(self, content):
		""" get the index of a specified row
			@param	content:	cell content whose row to get
			@return	integer index or None
		"""
		self.toRows()

		for row in self.spreadsheet:
			if content in row:
				return self.spreadsheet.index(row)

		return None

	def getSavePath(self):
		"""	Get the save path
			@return String of the save path
		"""
		return self.savePath

	def getSpreadsheet(self):
		"""	Get this Spreadsheet
			@return	List of rows and columns with content
		"""
		return self.spreadsheet

	def initialize(self, filePath = None, sep = "\t"):
		"""	Open the file and parse out rows and columns
			@param	filePath: spreadsheet file to load into memory
		"""
		openedFilePath	 = self.open(filePath).splitlines()

		for line in openedFilePath:
			self.spreadsheet.append(line.split(sep))

		self.filePath 	 = filePath
		self.loaded 	 = True
		self.toRows()

		if self.spreadsheet[0][0]=="columnName":
			del(self.spreadsheet[0])

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

	def open(self, filePath):
		"""	Opens an indicated text file for processing
			@param	filePath: path of file to load
			@return	String of opened text file
		"""
		fileIn1 = open(filePath, 'r')
		fileIn2 = fileIn1.read()
		fileIn1.close()
		return fileIn2

	def prepareForSave(self, spreadsheet=None, delimiter="\t"):
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
			row = [str(item) for item in row]
			append(delimiter.join(row))

		saveContent = "\n".join(rowList)

		return saveContent

	def refresh(self):
		""" Make sure all columns/rows are the same length
		"""
		self.transpose()
		self.transpose()
		self.initialized = True

	def rename(self, name=None, index=-1):
		"""	Rename a specified column
			@param	name: name for the column
			@param	index: column index
		"""
		# Transpose spreadsheet to edit column
		self.toColumns()

		self.spreadsheet[index][0] = name

	def removeColumn(self, column=-1):
		"""	remove a column in self.spreadsheet
			@param	column: index or string indicating column to remove
		"""
		self.toColumns()

		# if column variable is a string (i.e., column name), get column with the same name's index	
		if type(column) == type(str()):
			column = self.getColumnIndex(column)

		del self.spreadsheet[column]

	def removeRow(self, row=-1):
		"""	remove a row in self.spreadsheet
			@param	row: index of row to remove
		"""
		self.toRows()

		del self.spreadsheet[column]

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

	def save(self, savePath=None, saveContent=None, saveType='w', delimiter="\t"):
		"""	Write content out to a file
			@param	savePath: name of the file to be saved
			@param	saveContent: list of rows/columns to be saved
			@param	saveType: indicate overwrite ('w') or append ('a')
		"""
		if saveContent is None:
			saveContent = self.prepareForSave(delimiter=delimiter)
		else:
			saveContent = self.prepareForSave(spreadsheet=saveContent, delimiter=delimiter)

		saveFile = open(savePath, saveType)
		saveFile.write(saveContent)
		saveFile.close()

	def setColumn(self, column, newColumnCells, newHeader=False):
		"""	set a column to a new list of values
			@param	column: index or name of column
			@param	newColumnCells: list of new column values
		"""
		self.toColumns()
		if len(self.spreadsheet[self.getColumnIndex(column)]) <= 1:
			self.spreadsheet[self.getColumnIndex(column)].extend(newColumnCells)
		else:
			if newHeader == True:
				self.spreadsheet[self.getColumnIndex(column)] = newColumnCells
			else:
				self.spreadsheet[self.getColumnIndex(column)] = [self.spreadsheet[self.getColumnIndex(column)][0]]+newColumnCells

	def setCell(self, row, column, content):
		""" Set a cell to content
			@param	content: content to fill cell
		"""
		self.toRows()
		self.spreadsheet[row][column]=content

	def setData(self, data):
		""" set spreadsheet to new data
			@param	data: user specified spreadsheet data
		"""
		self.spreadsheet = data

	def setFilePath(self, filePath):
		"""	Set this object to a new file
			@param	filePath: location to new file
		"""
		self.initialize(filePath)

	def setSavePath(self, savePath):
		"""	Set the location for saved files
			@param	savePath: location to store saved files
		"""
		self.savePath = savePath

	def sort(self, column=0, reverse=False, hasTitle=True):
		""" Sort spreadsheet based on column
			@param	column: column to sort by
			@param	reverse: reverse sort
			@param	hasTitle: if True, sort spreadsheet from second row (i==1)
		"""
		self.toRows()
		
		column = self.getColumnIndex(column)

		if hasTitle == True:
			header = self.spreadsheet[0]
			body = sorted(self.spreadsheet[1:], key=lambda row:row[column], reverse=reverse)
			self.spreadsheet = [header] + body
		else:
			self.spreadsheet = sorted(self.spreadsheet, key=lambda row:row[column], reverse=reverse)

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

	def toString(self, fileToString = None):
		"""	Print input to screen
			@param fileToString: file to print out as string to screen.
		"""
		
		if fileToString is None:
			self.toRows()
			fileToString = self.spreadsheet
			
		if self.initialized:
			join = str.join

			# Loop through the lines to join them as strings
			for line in fileToString:

				line = [str(item).rjust(20, ' ') for item in line]
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

	p = Spreadsheet(columns=["col1", "booh", "bah", "ble"] + ["too", "me"])
	p.newColumn("blech")
	p.newColumn()
	print "P:\t", p.spreadsheet
	p.transpose()
	print "P:\t", p.spreadsheet
	p.transpose()
	print "P:\t", p.spreadsheet

	print p.prepareForSave()
	print p.addToColumn("col1",87)
	print p.addToColumn(0,55)
	print p.addToColumn(0,785)
	print p.addToColumn(1,3)
	p.refresh()
	print "FIRSTP:\t", p.spreadsheet
	p.setCell(3,2,"SETSELL")
	p.fillColumn("too", "can")
	p.addToColumn(4,"acd")
	print "P:\t", p.spreadsheet
	p.fillColumn("ble", "{0}, {1}", cellList=[("$A{0}", 0), ("$C${0}", 1)])
	p.fillColumn("col1", "{0}, {1}", cellList=[("{0}1", '0'), ("$C${0}", 1)])
	#p.fillColumn("ble", "{0}, {1}")#, cellList=[("$A{0}", 0), ("$C${0}", 1)])
	print "AFTER THE FILL:\t", p.spreadsheet
	p.toColumns()
	print "ddP:\t", p.spreadsheet
	p.removeColumn("bah")
	p.removeColumn("booh")
	print "Removed column:\t", p.spreadsheet
	print p.getColumnIndex('A')
	print p.getColumnIndex(00.0)
	#print p.sort(3)
	#print "P:\t", p.spreadsheet
	#print p.sort(4, hasTitle=False)
	#print "P:\t", p.spreadsheet
	#print "column index for 'A'", p.getColumnIndex('D')
