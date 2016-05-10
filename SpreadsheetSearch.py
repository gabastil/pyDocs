#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     SpreadsheetSearch.py
# Version:  1.0.0
# Author:   Glenn Abastillas
# Date:     October 22, 2015
#
# Purpose: Allows the user to:
#           1.) Compile a list of terms that correspond to sufficient (i.e., "-S") DICE code associated terms.
#           2.) Search insufficient results (i.e., "-I") in a language extract spreadsheet to search for possible missed cases.
#           3.) Add new columns to spreadsheet.
#           4.) Save output.
#
# This class does not have scripting code in place.
#
# This class is directly inherited by the following classes:
#   - ..\Analyze.py
#   - ..\DICESearch.py
# 
# Updates:
# 1. [2015/12/03] - added "savePath" variable to save() method.
# 2. [2015/12/04] - optimized processes.
# 3. [2015/12/07] - updated "format" function for loop in prepareTerms() method.
# 4. [2016/02/29] - changed wording of notes in line 17 from '... class is used in the following ...' to '... class is directly inherited by the following ...'.
# 5. [2016/02/29] - changed import statement from 'import File' to 'from File import Class' to allow for this class to inherit 'Class' instead of 'File.Class'. Version changed from 1.0.0 to 1.0.1.
# - - - - - - - - - - - - -
"""search a document for specific terms and create, manipulate, and save a spreadsheet containing the desired findings.

Search a specified document for specific terms as indicated by the DROOLs 
file. Presence of terms in the excerpts are noted in new columns created. 
The matching term and its priority index, an index that corresponds to its 
frequency in the DROOLs spreadsheet is also assigned to their own respective 
columns.

SpreadsheetSearch() extends SpreadsheetPlus() and DocumentPlus() and 
expands on them by analyzing their data against another spreadsheet 
containing DROOLS rules.
"""
__author__      = "Glenn Abastillas"
__copyright__   = "Copyright (c) October 22, 2015"
__credits__     = "Glenn Abastillas"

__license__     = "Free"
__version__     = "1.0.0"
__maintainer__  = "Glenn Abastillas"
__email__       = "a5rjqzz@mmm.com"
__status__      = "Deployed"

from SpreadsheetPlus    import SpreadsheetPlus
from DocumentPlus       import DocumentPlus

import time

class SpreadsheetSearch(SpreadsheetPlus, DocumentPlus):

	def __init__(self, fileForAnalysis = None, fileWithRules = None, columns=None, *DICECodes):
		"""	this class inherits attributes and methods from SpreadsheetPlus and
			DocumentPlus. This class enables the user to load two spreadsheets, 
			with the first containing raw data to be analyzed, and the second, 
			containing the parameters with which to analyze the data.

			Raw data is initialized and transformed for efficiency, removing 
			extraneous columns prior to processing the document.

			Three new columns are created and added to the spreadsheet: 

			1. "Results", which indicates whether or not a keyword was found in 
				the initially insufficient query.

			[DEPRECATED] 2. "Matched", which indicates the matched term found in the insuff-
				icient query.
			[DEPRECATED] 3. "Excerpt", which shows a snippet of the matched keyword in its 
				context.

			@param	fileForAnalysis : path to excerpts file to analyze
			@param	fileWithRules : path to Drools Rules spreadsheet
			@param	columns : columns to transform
			@param	*DICECodes : list of DICE Codes
		"""
		super(SpreadsheetSearch, self).__init__(fileForAnalysis, fileWithRules) # save paths for both spreadsheets

		if columns is None:
			columns = [1,5,6,7,8,9,10,11,3,0]

		if len(DICECodes) < 1:
			DICECodes = ["CH001", "CH002", "CH003", "CH004", "CH005", "CH006", "CH007", "CH008", "CH009", "CH010", \
						 "CH011", "CH012", "CH013", "CH014", "CH015", "CH016", "CH017", "CH018", "CH019", "CH020", \
						 "CH021", "CH022", "CH023", "CH024", "CH025", "CH026", "CH027", "CH028", "CH029", "CH030"]

		self.DICECodes = DICECodes         # List of DICE Codes to compile from the DROOLS RULES Spreadsheet
		self.stop_words = DocumentPlus().getStopWords()

		# Load and initialize both spreadsheets if indicated
		if fileForAnalysis is not None and fileWithRules is not None:
			#super(SpreadsheetSearch, self).load()                      		# load spreadsheets into memory
			super(SpreadsheetSearch, self).initialize(fileForAnalysis)            # intialize spreadsheets for processing (e.g., splitting on the comma)
			super(SpreadsheetSearch, self).transform(*columns)					# reduce spreadsheet columns to pertinent number

			#self.toColumns()
			#print "Spreadsheet length1:\t", len(self.spreadsheet)
			#super(SpreadsheetSearch, self).transform(8,4,5,6,0)	# reduce spreadsheet columns to pertinent number
			#print "Spreadsheet length2:\t", len(self.spreadsheet)

			self.newColumn("Results")	# [col index = 10] List that contains indication of and location of matched terms.
			self.newColumn("Indexes")	# [col index = 11] List that contains the matched term's index from prepare terms method.
			self.newColumn("Matched")	# [col index = 12] List that contains the matched term.
			#self.newColumn("Excerpt")	# [col index = 13] List that contains an excerpt of the matched term's context within a given scope.

		#print "Spreadsheet\t", self.spreadsheet[-1][:3], self.spreadsheet[-2][:3], self.spreadsheet[-3][:3], self.spreadsheet[-4][:3], self.spreadsheet[-5][:3]
		#print "Stopwords", self.stop_words[:10]

	def prepareTerms(self, dice, termIndex = 4, sufficiency = "-S"):
		"""	opens spreadsheet 2, which typically contains droolsrules.csv. 
			Users can extract associated terms to be used in searching the ex-
			cerpts document. Empty rows in the spreadsheet are skipped.
			@param	dice: DICE Code to compile
			@param	termIndex: column with terms. Column E (termIndex=4).
			@param	sufficiency: insufficient "-I" or sufficient "-S" terms.
			@return	list of stop-word-free terms to use for superFind
		"""
		termList = list()
		#termsToSearchFor = list()

		# CALL THESE JUST ONCE BEFORE LOOP(S)
		extend = termList.extend
		lower  = str.lower
		split  = str.split
		upper  = str.upper

		# Loop through lines in the Drools Rules spreadsheet
		for line in self.spreadsheetPlus:

			# Examine column E (index=5) if DICE + sufficiency match (e.g., CH001-S)
			if (dice + sufficiency in upper(line[2])) and line[termIndex] != "":
				extend(split(lower(line[termIndex])))
				## If column E (index=5) is empty, skip it
				#if line[termIndex] == "":
				#	pass

				# If column E has terms, add terms to term list
				#else:
				#	extend(split(lower(line[termIndex])))
		
		# Remove stop words and prepare output to count occurrence of terms
		output = list(set(termList))
		output = super(SpreadsheetSearch, self).removeStopWords(output)
		output = [[t, 0] for t in output]

		# CALL THESE JUST ONCE BEFORE LOOP(S)
		index = output.index

		# Loop through termList
		for t in termList:

			# Loop through output (set of terms) to add +1 for each time the term appears
			for o in output:

				# If the term is in the list, add +1
				if t == o[0]:
					output[index(o)][1] += 1

		return sorted(output, key=lambda tupleWithTermAndCount: tupleWithTermAndCount[1], reverse=True)

	def superFind(self, dice, termIndex = 4):
		"""	Takes a list of prepared terms with respect to the DICE code and 
			searches for those DICE associated terms in the excerpts spread-
			sheet. Users can analyze the resulting spreadsheet, which is tagged
			for appearance of associated terms and where the associated term was
			found - Left Column (Y-L), Right Column (Y-R), or Both (Y-LR).

			@param	dice : DICE Code requested/to search for
			@param	termIndex : location of the terms in Drools Rules

			@return	list of results
		"""
		terms = self.prepareTerms(dice = dice, termIndex = termIndex)       # get associated terms list
		
		# CALL THESE JUST ONCE BEFORE LOOP(S)
		getIndex = self.spreadsheet.index
		#extract = self.extract
		format  = str.format
		join    = str.join
		lower   = str.lower
		upper   = str.upper
		zfill   = str.zfill

		DICECodeRequested = dice
		self.toRows()

		#print DICECodeRequested
		# Loop through rows/lines to analyze excerpts
		#for line in self.spreadsheet:
		for i in range(len(self.spreadsheet))[1:]:
			#index = getIndex(line)        # assign the numerical index to this line
			row = self.spreadsheet[i]

			#DICECode = format("CH{0}-", zfill(str(int(row[0])), 3)) <-- turn back on after testing
			DICECode = row[0]

			# If there is text in this cell
			#if len(row[6]) > 0:
			if len(row[0]) > 0:
				DICECode = format("{0}", DICECode)
			else:
				DICECode = format("{0}", DICECode)

			# If the DICE Code in drools rules matches the requested DICE Code, continue analysis
			if DICECode==DICECodeRequested:
				#print DICECode, DICECodeRequested, DICECode==DICECodeRequested
				indexLeft = termIndex-1
				indexRight= termIndex+1

				termFound = False

				leadingText	= lower(row[indexLeft])
				trailingText= lower(row[indexRight])

				toIndexT = terms.index
				#toIndexL = leadingText.index
				#toIndexR = trailingText.index

				# Loop through the prepared terms
				for termTuple in terms:

					term = termTuple[0]

					# If the term is in both leading and trailing texts, assign "Y-LR",i.e., yes, left and right
					if (term in leadingText) and (term in trailingText):
						termFound = True
						code = "Y-LR"

					elif term in leadingText:
						termFound = True
						code = "Y-L"

					elif term in trailingText:
						termFound = True
						code = "Y-R"

					#print "termFound\t", termFound
					if termFound == True:
						#termTupleIndex = toIndexT(termTuple)
						#line = lower(join(' ', row[indexLeft:indexRight+1])).replace("|","").replace("  ", " ")
						spot = str(toIndexT(termTuple))
						#extract = self.extract(line, spot)

						#self.setCell(i, 10, code)
						#self.setCell(i, 11, excerpt)
						#self.setCell(i, 12, term)
						#self.setCell(i, 13, str(spot))
						#print code, '\t', extract[:50], '\t', "|", term
						self.setCell(i, 5, code)
						self.setCell(i, 6, spot)
						self.setCell(i, 7, term)
						#self.setCell(i, 8, extract)
						break
					else:
						#print "\t\tDoing N"
						#self.setCell(i,10, "N")
						self.setCell(i, 5, "N")
						#self.setCell(i, 6, "NA")
						#self.setCell(i, 7, "NA")

			#print "row", row
		self.toColumns()
		self.sort(0)
		#print self.spreadsheet[5]
		#print self.spreadsheet[6]
		#print self.spreadsheet[7]
		#print self.spreadsheet[8]
		return self.spreadsheet[5]

if __name__=="__main__":
	fileToAnalyze = u"M:\\DICE\\site - MaryWashington\\Extract2\\PHI\\samples\\mw2-ling-extracts.txt"
	fileWithRules = u".\\files\\droolsrules.txt"

	DICECodes = ["CH001", "CH002", "CH003", "CH004", "CH005", "CH006", "CH007", "CH008", "CH009", "CH010", \
				 "CH011", "CH012", "CH013", "CH014", "CH015", "CH016", "CH017", "CH018", "CH019", "CH020", \
				 "CH021", "CH022", "CH023", "CH024", "CH025", "CH026", "CH027", "CH028", "CH029", "CH030"]

	# Used for testing
	columns = [8,4,5,6,0]

	SS = SpreadsheetSearch(fileToAnalyze, fileWithRules, columns, *DICECodes)
	SS.superFind("CH001")
	print SS.prepareForSave()