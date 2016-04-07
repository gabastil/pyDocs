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
# 3. [2015/12/07] - updated "format" function for loop in prepareDICECode() method.
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


	def __init__(self, fileToAnalyze=None, pathToDROOLS=None, savePath=None):
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

		super(SpreadsheetSearch, self).__init__(filePath=fileToAnalyze, savePath=None) # Initialize Spreadsheet object

		self.DROOLS = SpreadsheetPlus(filePath=pathToDROOLS)				# Create DROOLs Spreadsheet

		# Load and initialize both spreadsheets if indicated
		if fileToAnalyze is not None and DROOLS is not None:
			super(SpreadsheetSearch, self).load()                      		# load spreadsheets into memory
			super(SpreadsheetSearch, self).initialize()                		# intialize spreadsheets for processing (e.g., splitting on the comma)
			super(SpreadsheetSearch, self).transform(1,5,6,7,8,9,10,11,3,0)	# reduce spreadsheet columns to pertinent number

			self.newColumn("Results", appendToTransformed=True)				# List that contains indication of and location of matched terms.
			self.newColumn("Term Index", appendToTransformed=True)				# List that contains the matched term's index from prepare terms method.
			self.newColumn("Matched Term", appendToTransformed=True)				# List that contains the matched term.
			self.newColumn("Excerpt", appendToTransformed=True)				# List that contains an excerpt of the matched term's context within a given scope.

			#-x self.results    = [" "] * len(self.spreadsheet[0])         		# List that contains indication of and location of matched terms.
			#-x self.indexes    = [" "] * len(self.spreadsheet[0])         		# List that contains the matched term's index from prepare terms method.
			#-x self.matched    = [" "] * len(self.spreadsheet[0])         		# List that contains the matched term.
			#-x self.excerpt    = [" "] * len(self.spreadsheet[0])         		# List that contains an excerpt of the matched term's context within a given scope.
		
		self.stop_words = super(SpreadsheetSearch, self).getStopWords()		# Removes extraneous, common words that do not contribute to analysis

		self.diceCodes = ["CH001", "CH002", "CH003", "CH004", "CH005",\
						  "CH006", "CH007", "CH008", "CH009", "CH010",\
						  "CH011", "CH012", "CH013", "CH014", "CH015",\
						  "CH016", "CH017", "CH018", "CH019", "CH020",\
						  "CH021", "CH022", "CH023", "CH024", "CH025",\
						  "CH026", "CH027", "CH028", "CH029", "CH030"]		# List of DICE Codes to compile from the DROOLS RULES Spreadsheet

	"""
	def addColumn(self, name = "New Column", fillWith = " ", sheet = 3, length = 0):
		"""	adds a new column to self.spreadsheet.
			@param name: indicate the name of the column to be added
			@param fillWith: indicate the filler to use for the blank column
			@param sheet: spreadsheet to add columns to (3 means both spreadsheets)
			@param length: length of column
		"""

		# If spreadsheets are not transposed, transpose them so loops loop over columns instead of rows
		if not self.spreadsheet_transposed:

			if len(self.spreadsheet) > 0:
				self.transpose(sheet)

		if len(self.spreadsheet) > 0:
			newColumn       = [fillWith] * len(self.spreadsheet[0])     # create a column sharing the same length as others in the spreadsheet
			newColumn[0]    = name                                      # assign a name to this column
			self.spreadsheet.append(newColumn)                          # add the column to the spreadsheet
		
		else:
			newColumn       = [fillWith] * length
			newColumn[0]    = name
			self.spreadsheet.append(newColumn)
			self.spreadsheet_transposed = True

	def fillColumn(self, name = None, fillWith = None, sheet = 1):
		""	inserts text into a specified column in the specified sheet.
			@param name: name of the column to fill
			@param fillWith: to insert into the cells of the specified column
			@param sheet: spreadsheet whose column will be filled with text
		""

		if name is None:
			return "No column specified. Please enter a name into the name variable (e.g., name = \"column\""

		# If spreadsheet is not transposed, tranpose it so loops loop over columns instead of rows
		if not self.spreadsheet_transposed:
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
			# If the spreadsheet is not transposed, revert it so that loops can work over columns rather than 
			# rows. 3 is indicated to tranpose both spreadsheets.
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
			self.transpose(sheet)

		# CALL THESE JUST ONCE BEFORE LOOP(S)
		lower   = str.lower
		replace = fillWith.replace
		# - - - - - - - - - - - - - - - - - -

		for column in self.spreadsheet:
			if lower(column[0]) == lower(name):
				for cell in range(len(column[1:])):
					#print "CELL RANGE", cell, "CELL TYPE", type(cell), "PLUS", cell+1
					#print "COLUMN", column[cell+1], fillWith, replace("{C}", str(cell+2))
					column[cell+1] = replace("{C}", str(cell+2))

	def consolidate(self):
		"	appends the results, matched, and excerpt columns to the existing 
			spreadsheet.
		"
		if not self.spreadsheet_transposed:
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
			# If the spreadsheet is not transposed, revert it so that loops can work over columns rather than 
			# rows. 3 is indicated to tranpose both spreadsheets.
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

			self.transpose(3)

		self.spreadsheet.append(self.results)       # append 'Results' column to spreadsheet
		self.spreadsheet.append(self.indexes)       # append 'Term Index' column to spreadsheet
		self.spreadsheet.append(self.matched)       # append 'Matched Term' column to spreadsheet
		#self.spreadsheet.append(self.excerpt)       # append 'Excerpt' column to spreadsheet

		# Use findLines from DocumentPlus for this
	def extract(self, text, center, scope = 50, upper = None):
		""	Get excerpts from text including leading and trailing text
			@param	text: String to analyze for excerpts
			@param	keyIndex: index position of key term
			@param	scope: number of leading and trailing characters
			@param	upper: return excerpt in upper case
			@return	String of excerpt
		""
		start = center - scope          # set the index for the beginning of the string
		end   = center + scope          # set the index for the end of the string
		
		if start < 0:
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
			# If the index for the beginning of the string is less than 0, change it to 0.
			# There are no negative indices.
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

			start = 0

		if end >= len(text):
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
			# If the index for the end of the string is extendes past the length of the string,
			# change it to the length of the string - 1 (because of 0-indexing).
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

			end = len(text) - 1

		if upper is not None:
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
			# If the user indicates a term for 'upper', then the excerpt will have a term in upper case.
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

			upperLen  = len(upper)
			newCenter = center + upperLen
			return text[start:center] + text[center:newCenter].upper() + text[newCenter:end]
		else:
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
			# If not, then excerpt will contain a term in lower case.
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

			upperLen = 0
			return text[start:end]
	"""

	def prepareDICECode(self, dice, termIndex=4, sufficiency="-S"):
		"""	Get a list of ASSOCIATED_TERMS and their frequencies
			@param	dice: DICE Code to get terms for
			@param	termIndex: column on DROOLs containing ASSOCIATED_TERMS
			@param	sufficiency: -S for sufficient or -I for insufficient
			@return	list of terms and their frequencies
		"""

		# CALL THESE JUST ONCE BEFORE LOOP(S)
		termList = list()
		extend	 = termList.extend

		lower	 = str.lower
		split	 = str.split
		upper	 = str.upper

		# Loop through each row in the DROOLs Rules spreadsheet
		for line in self.DROOLS:

			# If row matches specified DICE Code and is not empty, check contents of column E (i.e., 5th column, index 4)
			if (dice + sufficiency in upper(line[2])) and line[termIndex] != "":
				extend(split(lower(line[termIndex])))

		# Reduce termList to unique terms only without stop words and assign counter for next loop
		output = list(set(termList))
		output = DocumentPlus().removeStopWords(text=output)
		output = [[t, 0] for t in output]

		# CALL THESE JUST ONCE BEFORE LOOP(S)
		index = output.index

		# Loop through output list and count number of times term appears in term list
		for term, count in output:
			i = index([term,count])
			c = termList.count(term)

			output[i][1] = c

		return sorted(output, key = lambda x: x[1], reverse = True)
	
	def save(self, name = "SampleLingGlenn-2000-PROCESSED", saveAs = "txt", delimiter = "\t", savePath = None):
		"""	save the spreadsheet to a file. User can name the file and choose 
			the type of file to save as (i.e., txt, csv, tsv, etc.)

			name		-->  indicate the name of the file to output the data
			saveAs		-->  indicate the type of the file to output the data
			delimiter	-->  indicate the type of delimiter to use for the data
			savePath	-->  indicate the location where the save file should 
							 be stored
		"""

		name = "{0}{1}".format(str(time.strftime("%y%m%d_%H%M_")), name)            # (1) Append date and time stamp to name
		super(SpreadsheetSearch, self).save(name, saveAs, delimiter, savePath)      # (2) Save spreadsheet

		return name
		
	def superFind(self, dice, termIndex = 4, sufficiency = "-I", fileType = 1):
		"""	takes a list of prepared terms with respect to the DICE code and 
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
		if self.spreadsheet_transposed:
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
			# If the spreadsheet is transposed, revert it so that loops can work over rows rather than 
			# columns. 3 is indicated to tranpose both spreadsheets.
			#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

			self.transpose(3)
			
		terms = self.prepareDICECode(dice=dice, termIndex=termIndex)	# get associated terms list

		# REMOVE THESE LATER
		self.results[0] = "Results"										# add heading for 'Results' column
		self.indexes[0] = "Term Index"									# add heading for 'Term Index' column
		self.matched[0] = "Matched Term"								# add heading for 'Matched Term' column
		self.excerpt[0] = "Excerpt"										# add heading for 'Excerpt' column
		
		# CALL THESE JUST ONCE BEFORE LOOP(S)
		toIndex = self.spreadsheet.index
		extract = self.extract
		format  = str.format
		join    = str.join
		lower   = str.lower
		upper   = str.upper
		zfill   = str.zfill
		# - - - - - - - - - - - - - - - - - -

		# Loop through spreadsheet rows
		for line in self.spreadsheet:

			index = toIndex(line)        # assign the numerical index to this line

			# CALL THESE JUST ONCE BEFORE LOOP(S)
			#replace  = line.replace
			#toIndexL = line.index 
			# - - - - - - - - - - - - - - - - - -

			#print line

			if index == 0:
				pass
			else:
				if fileType == 1:
					DICECode = format("CH{0}-", zfill(str(int(line[0])), 3))
					if len(line[6]) > 0:
						DICECode = format("{0}S", DICECode)
					else:
						DICECode = format("{0}I", DICECode)
				else:
					DICECode = upper(line[0])


				if DICECode == dice + sufficiency:
					#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
					# If the DICE and sufficiency of this row match, then continue analysis. Otherwise, assign "NA",
					# meaning "not applicable", for this row.
					#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

					i = termIndex - 1
					j = termIndex + 1

					left  = lower(line[i])                 
					right = lower(line[j])  

					termFound = False                       
					
					# CALL THESE JUST ONCE BEFORE LOOP(S)
					toIndexT = terms.index
					toIndexL = left.index
					toIndexR = right.index
					# - - - - - - - - - - - - - - - - - -

					for term in terms:
					#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
					# Cycle through the terms.
					#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

						if (term[0] in left) and (term[0] in right):
						#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
						# If the term is located in both left and right columns, assign "Y - LR" meaning "Yes - Left 
						# and Right" to the results column.
						#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

							self.results[index] = "Y - LR"
							termFound = True
							
							#print "TERM", term, term[0] in left, line

							line    = lower(join(' ', line[i:j+1])).replace("|","").replace("  ", " ")
							spot    = toIndexL(term[0])
							
							self.excerpt[index] = extract(line, spot, upper = term[0]) # get the context extract for this row
							self.matched[index] = term[0]
							self.indexes[index] = str(toIndexT(term))
							
							break   # Terms found in both pre- and post-text already, break the loop and move on.

						elif term[0] in left:
						#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
						# If the term is located in the left column, assign "Y - L" meaning "Yes - Left" to the results
						# column.
						#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

							self.results[index] = "Y - L"
							termFound = True
							
							line    = lower(join(' ', line[i:j+1])).replace("|","").replace("  ", " ")
							spot    = toIndexL(term[0])
							
							self.excerpt[index] = extract(line, spot, upper = term[0]) # get the context extract for this row
							self.matched[index] = term[0]
							self.indexes[index] = str(toIndexT(term))

							break   # Terms found in both pre- or post-text already, break the loop and move on.

						elif term[0] in right:
						#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
						# If the term is located in the right column, assign "Y - R" meaning "Yes - Right" to the results
						# column.
						#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

							self.results[index] = "Y - R"
							termFound = True
							
							line    = lower(join(' ', line[i:j+1])).replace("|","").replace("  ", " ")
							spot    = toIndexR(term[0])
							
							self.excerpt[index] = extract(line, spot, upper = term[0]) # get the context extract for this row
							self.matched[index] = term[0]
							self.indexes[index] = str(toIndexT(term))

							break   # Terms found in both pre- or post-text already, break the loop and move on.

					if termFound == False:
					#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
					# If the term was not found in this line, assign "N" to the results column.
					#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

						self.results[index] = "N"

		return self.results
