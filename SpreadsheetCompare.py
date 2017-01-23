# Name: SpreadsheetCompare.py
# Version: 1.0
#   (updated: 2015/12/09: change some loops to generators. Edited the introduction comment to show where SpreadsheetPlus is inherited.)
# Author: Glenn Abastillas
# Date: 10/15/2015
# Purpose: Allows the user to:
#           1.) Load a variants db text file (obtain from Access odbc)
#           2.) Check if variant(s) found in text to compare
#           3.) Add new columns to spreadsheet.
#           4.) Save output.
#
# To see the script run, go to the bottom of this page.
#
# This class is used in the following classes:
#
# - - - - - - - - - - - - -
import SpreadsheetPlus

class SpreadsheetCompare(SpreadsheetPlus.SpreadsheetPlus):

    def __init__(self, f1 = None, f2 = None, compareStr = None):
        """
            This class inherits attributes and methods from SpreadsheetPlus.
            This class enables the user to load two spreadsheets, with the first containing an up-to-date database of variant keywords,
            and the second, containing the data to be compared.

            Three new columns are created and added to the spreadsheet: 1. "Results", which indicates whether or not a keyword was found in the initially insufficient query.
                                                                        2. "Matched", which indicates the matched term found in the insufficient query.
                                                                        3. "Excerpt", which shows a snippet of the matched keyword in its context.
            Stop words are drawn from the DocumentPlus class.
        """

        super(SpreadsheetCompare, self).__init__(f1, f2)
        super(SpreadsheetCompare, self).load()
        super(SpreadsheetCompare, self).initialize()

        if "var" in f2.lower():
            self.db = set([item[0].replace("\"", "") for item in self.spreadsheet2])
        else:
            self.db = set([item[0].replace("\"", "") for item in self.spreadsheet])
            
    def addColumn(self, name = "New Column", fillWith = " "):
        """
            Adds a new column in the spreadsheet.

            name:       indicate the name of the column to be added
            fillWith:   indicate the filler to use for the blank column
        """

        if not self.spreadsheet_transposed:
            #=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#
            # If the spreadsheet is not transposed, revert it so that loops can work over columns rather than 
            # rows. 3 is indicated to tranpose both spreadsheets.
            #=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*#

            self.transpose(3)

        newColumn = [fillWith] * len(self.spreadsheet[0])       # create a column sharing the same length as others in the spreadsheet
        newColumn[0] = name                                     # assign a name to this column
        self.spreadsheet.append(newColumn)  

    def inDataBase(self, compareStr = None):
        if compareStr == None:
            return "Missing string to compare."

        print compareStr in self.db
                    
if __name__ == "__main__":
    """
        This will run if the script is run by itself.
    """


    variantsFile = "files\\20151015_Variants.txt"
    excerptsFile = "files\\Sample-ling-excerpts-all-kaleida-Glenn-2000-PROCESSED.csv"
    
    outputFileName = "Sample-ling-excerpts-all-kaleida-Glenn-2000-db_comparison_done"

    s = SpreadsheetCompare(variantsFile, excerptsFile)
    
    print s.inDataBase(compareStr = "rad")
