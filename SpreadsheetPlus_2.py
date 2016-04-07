#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     SpreadsheetPlus.py
# Version:  1.2.0
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
__version__     = "1.2.1"
__maintainer__  = "Glenn Abastillas"
__email__       = "a5rjqzz@mmm.com"
__status__      = "Deployed"

from Spreadsheet import Spreadsheet
import os

class SpreadsheetPlus(Spreadsheet):

	def __init__(self, filePath = None, savePath = None):
		""" Initialize an instance of this class
			@param  filePath: path of the first  spreadsheet file
			@param  savePath: write to this location
		"""
		self.transformedSpreadsheet = list()	# Contains transformed spreadsheet
		self.transformed = False				# Checks if self.spreadsheet was transformed

		# If filePath is specified, initialize it in the parent class Spreadsheet
		if filePath is not None:
			super(SpreadsheetPlus, self).__init__(filePath=filePath, savePath=savePath)

	# [USED TO BE consolidate() in SpreadsheetSearch]
	def appendColumn(self, *columns):
		"""	Append column(s) specified to spreadsheet
		"""
		# Transpose spreadsheet to loop over columns
		self.spreadsheetToColumns()

		# Loop through input columns to add to spreadsheet
		for column in columns:
			self.spreadsheet.append(column)

	# [USED TO BE addColumn() in SpreadsheetSearch]
	def newColumn(self, name="New Column", fillWith=" ", appendToTransformed=False):
		"""	adds a new column to specified spreadsheet
			@param name: name of the column
			@param fillWith: filler for the blank column
			@param appendToTransformed: add column to transformedSpreadsheet if True
		"""
		# Transpose spreadsheet to loop over columns
		self.spreadsheetToColumns()

		newColumn	 = [fillWith] * len(self.spreadsheet[0])	# create a column sharing the same length as others in the spreadsheet
		newColumn[0] = name										# assign a name to this column

		if appendToTransformed:
			self.transformedSpreadsheet.append(newColumn)		# Add to transformedSpreadsheet
		else:
			self.spreadsheet.append(newColumn)					# add the column to the spreadsheet

	def fillColumn(self, name=None, fillWith=" "):
		"""	inserts text into a specified column in the specified sheet
			@param name: name of the column to fill
			@param fillWith: to insert into the cells of the specified column
		"""
		# If name is None, return error
		if name is None:
			return ValueError("No column specified. Please enter a name into the name variable (e.g., name = \"column\")")

		# Transpose spreadsheet to loop over columns
		self.spreadsheetToColumns()

		# CALL THESE JUST ONCE BEFORE LOOP(S)
		lower = str.lower

		# Loop through the columns of this spreadsheet until the correct column is found
		for column in self.spreadsheet:

			# If the input name matches the column name, loop through the rows of this columns
			if lower(column[0]) == lower(name):

				# Loop through the rows of this column
				for cell in column[1:]:
					cell = fillWith

				break

	def spreadsheetToColumns(self):
		"""	Transpose spreadsheet so loops work over columns instead of rows
		"""
		if not self.transposed:
			self.transpose()

	def transform(self,*newColumns):
		"""	Create a new spreadsheet consisting of columns specified
			@param	newColumns:	a list of column indices (args)
		"""
		# Return self.spreadsheet to original form (i.e., not transposed)
		if self.transposed:
			self.transpose()

		# If no columns are specified in arguments use default
		if len(newColumns) < 1:

			#Corresponds to columns indicating: DICE-Code, pre-text, match, and post-text
			newColumns = [1,4,5,6] 

		else:
			newColumns = list(newColumns)
			
		# CALL THESE JUST ONCE BEFORE LOOP(S)
		newSpreadsheet = list()
		append = newSpreadsheet.append

		for column in newColumns:
			append(self.spreadsheet[column])

		transformedSpreadsheet = newSpreadsheet
		self.transformed = True

	def getSpreadsheet(self, transformed=False):
		""" Get this spreadsheet or transformed spreadsheet
			@param	transformed: indicated spreadsheet to retrieve
			@param	newColumns:	a list of column indices (args)
			@return	List of rows and columns with content
		"""
		# If spreadsheet is transformed, return transformedSpreadsheet
		if transformed and self.transformed:
			return self.transformedSpreadsheet

		# Return regular spreadsheet otherwise
		else:
			return self.spreadsheet

if __name__=="__main__":

	filePath1 = "../files/debridementSamples.txt"

	d = SpreadsheetPlus(filePath1)
	print len(d.getSpreadsheet()), d.getSpreadsheet()
	d.newColumn("new test", "This is a test")
	print len(d.getSpreadsheet()), d.getSpreadsheet()[-1]
	d.appendColumn(["this", "is", "a", "test"])
	print len(d.getSpreadsheet()), d.getSpreadsheet()[-1]