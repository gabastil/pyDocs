#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     SpreadsheetSearch.py
# Version:  1.4.0
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
__version__     = "1.4.0"
__maintainer__  = "Glenn Abastillas"
__email__       = "a5rjqzz@mmm.com"
__status__      = "Deployed"

from SpreadsheetPlus    import SpreadsheetPlus
from DocumentPlus       import DocumentPlus

import time

class SpreadsheetSearch(SpreadsheetPlus, DocumentPlus):

	def __init__(self, fileToAnalyze = None, fileWithRules = None, *DICECodes):
		"""	this class inherits attributes and methods from SpreadsheetPlus and
			DocumentPlus. This class enables the user to load two spreadsheets, 
			with the first containing raw data to be analyzed, and the second, 
			containing the parameters with which to analyze the data.

			Raw data is initialized and transformed for efficiency, removing 
			extraneous columns prior to processing the document.

			Three new columns are created and added to the spreadsheet: 

			1. "Results", which indicates whether or not a keyword was found in 
				the initially insufficient query.
			2. "Matched", which indicates the matched term found in the insuff-
				icient query.
			3. "Excerpt", which shows a snippet of the matched keyword in its 
				context.

			Stop words are drawn from the DocumentPlus class.
		"""

		super(SpreadsheetSearch, self).__init__(fileToAnalyze, fileWithRules) # save paths for both spreadsheets

		# Load and initialize both spreadsheets if indicated
		if fileToAnalyze is not None and fileWithRules is not None:
			#super(SpreadsheetSearch, self).load()                      		# load spreadsheets into memory
			super(SpreadsheetSearch, self).initialize(fileToAnalyze)            # intialize spreadsheets for processing (e.g., splitting on the comma)
			#super(SpreadsheetSearch, self).transform(1,5,6,7,8,9,10,11,3,0)	# reduce spreadsheet columns to pertinent number
			#self.toColumns()
			#print "Spreadsheet length1:\t", len(self.spreadsheet)
			super(SpreadsheetSearch, self).transform(8,4,5,6,0)	# reduce spreadsheet columns to pertinent number
			#print "Spreadsheet length2:\t", len(self.spreadsheet)

			self.newColumn("Results")	# [col index = 10] List that contains indication of and location of matched terms.
			self.newColumn("Indexes")	# [col index = 11] List that contains the matched term's index from prepare terms method.
			self.newColumn("Matched")	# [col index = 12] List that contains the matched term.
			self.newColumn("Excerpt")	# [col index = 13] List that contains an excerpt of the matched term's context within a given scope.

		#self.stop_words = super(SpreadsheetSearch, self).getStopWords()		# Removes extraneous, common words that do not contribute to analysis

		self.DICECodes = DICECodes         # List of DICE Codes to compile from the DROOLS RULES Spreadsheet
		self.stop_words = DocumentPlus().getStopWords()
		print "Spreadsheet\t", self.spreadsheet[-1][:3], self.spreadsheet[-2][:3], self.spreadsheet[-3][:3], self.spreadsheet[-4][:3], self.spreadsheet[-5][:3]
		print "Stopwords", self.stop_words[:10]

	def extract(self, text, center, scope = 50, upper = None):
		"""	extract allows users to grab an excerpt of the indicated text via 
			'center'. Users can grab an extract to help in analyzing context 
			for the chosen keyword match.
			@param	text: string to be analyzed
			@param	center: position of matched term
			@param	scope: leading and trailing characters to include
			@param	upper: return as upper case
			@return	String of the excerpt 
		"""
		start = center - scope	# set the index for the beginning of the string
		end   = center + scope	# set the index for the end of the string
		
		# Negative indices become 0
		if start < 0:
			start = 0

		# Indices greater than the length of the text are given the length of the text - 1
		if end >= len(text):
			end = len(text) - 1

		# Indicate upper case or not
		if upper is not None:
			upperLen  = len(upper)
			newCenter = center + upperLen
			return text[start:center] + text[center:newCenter].upper() + text[newCenter:end]

		else:
			upperLen = 0
			return text[start:end]

	def prepareTerms(self, dice, termIndex = 4, sufficiency = "-S"):
		"""	opens spreadsheet 2, which typically contains droolsrules.csv. 
			Users can extract associated terms to be used in searching the ex-
			cerpts document. Empty rows in the spreadsheet are skipped.
			@param	dice: DICE Code to compile
			@param	termIndex: column with terms. Column E (termIndex=4).
			@param	sufficiency: insufficient "-I" or sufficient "-S" terms.
			@return	list of stop-word-free terms to use for superFind
		"""
		termList         = list()
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

	def superFind(self, dice, termIndex = 4, fileType = 1):
		"""	Takes a list of prepared terms with respect to the DICE code and 
			searches for those DICE associated terms in the excerpts spread-
			sheet. Users can analyze the resulting spreadsheet, which is tagged
			for appearance of associated terms and where the associated term was
			found - Left Column (Y- L), Right Column (Y - R), or Both (Y - LR).

			dice 		--> indicate the DICE code you are interested in com-
							piling (e.g., CH001).
			termIndex 	--> column in the spreadsheet (0-index) that contains
							associated terms. default is 4, i.e., column E.
			sufficiency --> indicate whether you want to compile associated 
							terms belonging to insufficient "-I" or suffi-
							cient "-S" terms.

			Returns a list of results from the find.
		"""
		terms = self.prepareTerms(dice = dice, termIndex = termIndex)       # get associated terms list
		
		# CALL THESE JUST ONCE BEFORE LOOP(S)
		getIndex = self.spreadsheet.index
		extract = self.extract
		format  = str.format
		join    = str.join
		lower   = str.lower
		upper   = str.upper
		zfill   = str.zfill

		DICECodeRequested = dice
		self.toRows()

		print DICECodeRequested
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
				toIndexL = leadingText.index
				toIndexR = trailingText.index

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

					if termFound == True:
						termTupleIndex = toIndexT(termTuple)
						line = lower(join(' ', row[indexLeft:indexRight+1])).replace("|","").replace("  ", " ")
						spot = toIndexL(termTuple[0])
						extract = self.extract(line, spot, upper=term)

						#self.setCell(i, 10, code)
						#self.setCell(i, 11, excerpt)
						#self.setCell(i, 12, term)
						#self.setCell(i, 13, str(spot))
						print code, extract, "||||", term
						self.setCell(i, 5, code)
						self.setCell(i, 6, extract)
						self.setCell(i, 7, term)
						self.setCell(i, 8, str(spot))
						break
					else:
						#self.setCell(i,10, "N")
						self.setCell(i,5, "N")

			#print "row", row
		self.toColumns()
		print self.spreadsheet[5]
		return self.spreadsheet[5]

if __name__=="__main__":
	fileToAnalyze = u"M:\\DICE\\site - MaryWashington\\Extract2\\PHI\\samples\\mw2-ling-extracts.txt"
	fileWithRules = u"C:\\Users\\a5rjqzz\\Desktop\\Python\\pyDocs\\files\\droolsrules.txt"

	DICECodes = ["CH001", "CH002", "CH003", "CH004", "CH005", "CH006", "CH007", "CH008", "CH009", "CH010", \
				 "CH011", "CH012", "CH013", "CH014", "CH015", "CH016", "CH017", "CH018", "CH019", "CH020", \
				 "CH021", "CH022", "CH023", "CH024", "CH025", "CH026", "CH027", "CH028", "CH029", "CH030"]

	SS = SpreadsheetSearch(fileToAnalyze, fileWithRules, *DICECodes)
	SS.superFind("CH003")