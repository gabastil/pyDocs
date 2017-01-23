#=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
The folder contains the following classes:

BATCHREADER (inherits: OBJECT)
	# BatchReader converts a multiple *.rtf files into *.txt and places them in a 'converted' folder.
	# This class makes use of Rtf15Reader and PlaintextWriter from the pyth plugin downloaded here: https://pypi.python.org/pypi/pyth/0.5.4
	# There are no methods or functions in this class. All processes run automatically upon instantiation.

	input:		filePath, i.e., folder location of *.rtf files

	processes:	None

DOCUMENT (11/24/2015 - Make DEPRECATED)
	"Document is a class that allows for a simple search for a keyword in a document. Users can load the document (.txt), and search for a simple string."

	input:		filepath, i.e., 1 document
				savepath

	processes:	
				find()			1. searches for search term
								2. saves search term, if found, and set words before and after the search term
								3. returns list of strings

				load()			1. loads document into memory

								n. arguments:
									filePath: path containing document(s)

				reset()			1. clears all data

				save()			1. saves text file to output file

								n. arguments:
									name: name of the output text file

				setSavePath()	1. change path for output file

								n. arguments:
									savePath: path to be set

				toString()		1. prints argument to screen
								2. prints spreadsheet to screen if no argument present

								n. arguments:
									textType: 'text' (strings) or 'data' (tables)

									fileToString: object to be printed to screen

DOCUMENTPLUS (inherits: DOCUMENT)
	"DocumentPlus is a class that allows for the removal of stop words and punctuation from a term. User can clean text data by removed stop words and punctuation, resulting in an easier text document to analyze."

	input:		filepath, i.e., 1 document
				savepath

	processes:	
				find()			1. searches for search term
								2. saves search term, if found, and set words before and after the search term
								3. returns list of strings

				getStopWords()	1. returns a list of stop words.

				remove_stop_words()	1. remove stop words from the term

									n. arguments:
										term: string term with stop words to be removed

										splitBy: character to split text string by

										asString: returns a string join by 'splitBy' character

				remove_punctuation() 1. remove punctuation from term

									n. arguments:
										term: string term with stop words to be removed

										asString: returns a string join by 'splitBy' character

INFOTEXT (base class)

LOOPDIR (base class)

SPREADSHEET (base class)
	"Spreadsheet is a class that allows for simple manipulation of a spreadsheet file. Users can load the spreadsheet into memory, transpose the spreadsheet, and search for simple strings."

	input:		1 spreadsheet (e.g., .csv, .tsv)
				does not accept Excel files

	processes:	
				find()			1. searches for search term
								2. returns list of tuples: (term, col, row)

								n. arguments:
									searchTerm: term to be searched in self.spreadsheet

				initialize()	1. parses rows into cells (req.: load())

				load()			1. loads spreadsheet into memory
								2. parses into rows

								n. arguments:
									spreadsheet_file: spreadsheet file to load (e.g., spreadsheet.csv). does not accept Excel files.

				toString()		1. prints argument to screen
								2. prints spreadsheet to screen if no argument present

								n. arguments:
									fileToString: object to be printed to screen


				transpose()		1. makes 'rows' in self.spreadsheet into 'columns'
								2. makes 'columns' in self.spreadsheet into 'rows'

				determineGroupByRange()	1. empty

										n. arguments:
											columnToBeDetermined:
											rangeToBeDetermined:

SPREADSHEETPLUS (inherits: SPREADSHEET)
	"SpreadsheetPlus is a class that allows for manipulation of two spreadsheets - the 'excerpts' spreadsheet and the 'DICE codes' spreadsheet. Users can load the spreadsheet into memory, tranpose the spreadsheet, transform the excerpts spreadsheet by creating a new spreadsheets with columns specified, search for simple strings, and save the new spreadsheet into a file specified by the user."

	input:		2 spreadsheets (e.g., .csv, .tsv)
					-	1) spreadsheet with excerpts
					-	2) spreadsheet with DICE codes

				does not accept Excel files

	processes:	
				find()			1. searches for search term
								2. returns list of tuples: (term, col, row)

								n. arguments:
									searchTerm: term to be searched in self.spreadsheet

				initialize()	1. parses rows into cells (req.: load())


				load()			1. loads spreadsheet into memory
								2. parses into rows

								n. arguments:
									spreadsheet_file: spreadsheet with excerpts.

									spreadsheet_file2: spreadsheet with DICE codes.

				save()			1. save transformed spreadsheet to drive.

								n. arguments:
									name: 		name of file to be created

									saveAs: 	type of file to be created
												default is '.txt'

									delimiter: 	type of delimiter to separate columns in file
												default is '\t'

				toString()		1. prints argument to screen
								2. prints spreadsheet to screen if no argument present

								n. arguments:
									fileToString: object to be printed to screen

				transform()		1. creates a new spreadsheet from the excerpts spreadsheet

								n. arguments:
									*newColumns: specify columns to be included in the new spreadsheet.


				transpose()		1. makes 'rows' in self.spreadsheet into 'columns'
								2. makes 'columns' in self.spreadsheet into 'rows'

								n. arguments:
									sheet: indicates which spreadsheet to transpose. '1' to transpose first or 'excerpts' spreadsheet. '2' to transpose second or 'DICE codes' spreadsheet.

SPREADSHEETSEARCH (inherits: SPREADSHEETPLUS, DOCUMENTPLUS)
	"SpreadsheetSearch is a class that allows the user to search an input spreadsheet (e.g., linguistic excerpts extracted from deidentified text files) for matching terms from another spreadsheet containing keywords and rules (e.g., DROOLS rules). A new spreadsheet is returned containing the pertinent columns to perform an engine tuning analysis."

	input:		2 spreadsheets (e.g., .csv, .tsv)
					-	1) spreadsheet with excerpts
					-	2) spreadsheet with DICE codes

				does not accept Excel files

	processes:	
				prepareTerms()	1. opens up the droolsrules.csv spreadsheet
								2. extracts associated terms connected to provided DICE codes
								3. returns a sorted list of associated terms

								n. arguments:
									dice:			indicate the DICE code you are interested in compiling (e.g., CH001)

									termIndex:		indicate column in the spreadsheet (0-index) that contains associated terms
													default is 4 (i.e., column E)

									sufficiency:	indicate whether you want to compile associated terms belonging to insufficient "-I" or sufficient "-S" terms
													default is "-S"

				superFind()		1. get list of associated terms for search
								2. cycle through rows in spreadsheet to search for associated terms
								3. assign labels for each row in the 'results' column 
									a. (values: "NA" 		- not applicable, 
												"N" 		- no, 
												"Y - LR" 	- yes in left and right, 
												"Y - L" 	- yes in left)
								4. returns a list of results

								n. arguments:
									dice:			indicate the DICE code you are interested in compiling (e.g., CH001)

									termIndex:		indicate column in the spreadsheet (0-index) that contains associated terms
													default is 4 (i.e., column E)

									sufficiency:	indicate whether you want to compile associated terms belonging to insufficient "-I" or sufficient "-S" terms
													default is "-I"

				extract()		1. get an extract around the associated term
								2. return the extract with upper case text if text is entered into the upper parameter

								n. arguments:
						            text:       	string to be analyzed for excerpts

						            center:     	indicates position of matched term

						            scope:      	how many characters in front of and behind the matched term to include
						                        	default is 50

						            upper:      	indicates whether or not to return the excerpt in UPPER case.
						                        	default is None

				consolidate()	1. add the 'results', 'matched term', and 'excerpt' columns to the spreadsheet

				addColumn()		1. add a new column to the spreadsheet

								n. arguments:
									name:			indicate the name of the column to be added

									fillWith:		indicate the filler to use for the blank column

				save()			1. save processed spreadsheet to drive

								n. arguments:
						            name:       	indicate the name of the file to output the data
						            
						            saveAs:     	indicate the type of the file to output the data
						            
						            delimiter:  	indicate the type of delimiter to use for the data

	use:		1. load an excerpt(s) spreadsheet and droolsrules spreadsheet
					s = SpreadsheetSearch(excerptsSpreadsheet, 		droolsRulesSpreadsheet)
								also ... (f1 = excerptsSpreadsheet, f2 = droolsRulesSpreadsheet)

				2. cycle through DICE codes and call superFind(code)
					for code in diceCodes:
						s.superFind(code)

				3. consolidate new results column with the spreadsheet
					s.consolidate()

				4. add additional columns as needed
					s.addColumn("eval")

				5. save output to file if needed
					s.save("NameOfOutput")

SPREADSHEETSEARCHLOG (base class)


