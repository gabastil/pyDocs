# Name: CrosscheckDictionary.py
# Version: 1.0.0 
# Author: Glenn Abastillas
# Date: February 9, 2016
# Purpose: Allows the user to:
#           1.) Check found terms' existence in dictionary (pulled from W0157340 SQL Database)
#           2.) Add a new column stating whether or not the term was found
# To see the script run, go to the bottom of this page. 
# - - - - - - - - - - - - -
"""	compare terms/phrases in Column P [index=15] of a spreadsheet to those in the DICE Dictionary for duplicates.

CrosscheckDictionary compares terms/phrases in question in Column Q [index = 16] of the YYMMDD_HHMM_ling-excerpts-CLIENT-processed_B.tsv spreadsheet (as *.txt) to terms/phrases present in the DICE Dictionary found in the SQL databased in workstation W0157340. 

This class uses the Spreadsheet class to load and initialize the spreadsheet file and the dictionary file for comparison. To compare, the column containing terms in both of these files' spreadsheets are extracted and compared. If there are term matches in the dictionary, "IN DICT" is inserted in a newly appended column of the Spreadsheet object. If there is no match, "NOT IN DICT" is inserted. 
"""
from Spreadsheet 		  import Spreadsheet

import os
import time

class CrosscheckDictionary(object):

	EMPTY_CELL  = "EMPTY CELL"
	IN_DICT 	= "IN DICT"
	NOT_IN_DICT = "NOT IN DICT"

	def __init__(self, spreadsheetFile, dictionaryFile, spreadsheetIndex = 0, dictionaryIndex = 0):
		"""	initialize this object by extracting columns from the spreadsheet
			and the dictionary files specified.

			@param	spreadsheetFile	 - path to spreadsheet
			@param	dictionaryFile	 - path to dictionary
			@param	spreadsheetIndex - index of columns to be checked
			@param	dictionaryIndex	 - index of terms to check against
		"""
		# Initialize these variables using the input parameters
		self.spreadsheet, self.spreadsheetIndex, self.dictionaryIndex = self.open(spreadsheetFile, dictionaryFile, spreadsheetIndex, dictionaryIndex)

	def open(self, spreadsheetFile, dictionaryFile, spreadsheetIndex = 0, dictionaryIndex = 0):
		"""	open the spreadsheet file and dictionary file and extract the
			specified columns for comparison.

			@param spreadsheetFile	- path to spreadsheet file
			@param dictionaryFile	- path to dictionary file
			@param spreadsheetIndex	- index of terms in spreadsheet
			@param dictionaryIndex	- index of terms in dictionary

			@return	tuple (Spreadsheet object, spreadsheet, dictionary)
		"""
		# Create and initialize Spreadsheet objects for the spreadsheet file and dictionary file
		files = [Spreadsheet(f) for f in [spreadsheetFile, dictionaryFile]]

		for s in files:
			s.load()
			s.initialize()

		# Make both spreadsheet (generator) and dictionary (list) for the check method
		spreadsheet = (line[spreadsheetIndex].lower() for line in files[0].spreadsheet[1:])
		dictionary  = [line[dictionaryIndex].lower()  for line in files[1].spreadsheet if len(line[dictionaryIndex]) > 1]

		return files[0], spreadsheet, dictionary

	def check(self, SpreadsheetObject, spreadsheet, dictionary):
		"""	compare the spreadsheet list and the dictionary list
			for any matches and record matches in a new column.

			@param	SpreadsheetObject - Spreadsheet object to add new column to
			@param	spreadsheet - list of terms from spreadsheet
			@param	dictionary - list of terms from dictionary

			@return	Spreadsheet object with new column added
		"""
		# Set new spreadsheet header
		SpreadsheetObject.spreadsheet[0] += ["[C] In DICE Dict?"]

		# Initiate index to 1 to skip spreadsheet header
		index = 1

		# Loop through the spreadsheet to be check and check term against the dictionary if it is not blank
		for row in spreadsheet:

			# If the row is blank then say "EMPTY CELL"
			if row.isspace():
				SpreadsheetObject.spreadsheet[index] += [self.EMPTY_CELL]

			# If the row is not blank check if term in row is in the dictionary
			else:
				#print("Checking row: {0}".format(row))
				# If the term is in the dictionary, say "IN DICT"
				if row in dictionary:
					SpreadsheetObject.spreadsheet[index] += [self.IN_DICT]
					print("Term in dictionary: {0}".format(row))	

				# If the term is not in the dictionary, say "NOT IN DICT"
				else:
					SpreadsheetObject.spreadsheet[index] += [self.NOT_IN_DICT]

			# Increment index
			index += 1

		return SpreadsheetObject

	def save(self, spreadsheet, outputPath):
		"""	make spreadsheet (list of lists) into a string and save to path.

			@param	spreadsheet: list containing rows as list
			@param	outputPath: path to save directory
		"""
		# Convert list to string for saving
		spreadsheetColumnsAdded  = ('\t'.join(line) for line in spreadsheet)
		spreadsheetNewLinesAdded = '\n'.join(spreadsheetColumnsAdded)

		# Save string to file specified in path
		writeOut = open(outputPath, 'w')
		writeOut.write(spreadsheetNewLinesAdded)
		writeOut.close()

if __name__=="__main__":
	""" run as a script if this file is run as a stand-alone program
	"""

	# Path to spreadsheet file and dictionary files
	spreadsheetFile = "C:\\Users\\a5rjqzz\\Desktop\\Excel\\20160318 tanner\\test.txt"
	dictionaryFile  = "C:\\Users\\a5rjqzz\\Documents\\Variants\\20160321_variants.txt"

	# Index to columns in the spreadsheet and dictionary to be compared
	spreadsheetFileColumnIndex = 0
	dictionaryFileColumnIndex  = 0

	# Create CrosscheckDictionary object
	cd = CrosscheckDictionary(spreadsheetFile, dictionaryFile, spreadsheetFileColumnIndex, dictionaryFileColumnIndex)

	# Check and then save run
	cd.check(cd.spreadsheet, cd.spreadsheetIndex, cd.dictionaryIndex)
	cd.save(cd.spreadsheet, spreadsheetFile)