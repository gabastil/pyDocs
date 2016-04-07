# Name: Analyze.py
# Version: 1.2 
# Author: Glenn Abastillas
# Date: December 3, 2015
# Purpose: Allows the user to:
#           1.) Analyze sample excerpts in a folder.
#           2.) Analyze the results against a DROOLS rules spreadsheet.
#
# To see the script run, go to the bottom of this page.
#
# This class is used in the following classes:
# - - - - - - - - - - - - -

from collections import defaultdict
import time, re, os

class LoadDB(object):

	def __init__(self):
		self.dbPath = "C:\\Users\\a5rjqzz\\Documents\\search-keywords-all.txt" #20151027_Variants.txt"
		#self.dbPath = "M:\\DICE\\Trinity\\Extracts\\stpeters\\St Peters eMD Re-Tune Doc Extracr\\Extract1\\samples\\Sample_StPeters_Glenn-2000-deid\\cln\\200001.txt"
		self.db = defaultdict(set)
		self.dbSet = set()
		self.dbList = list()
		self.dbDICE = list()
		self.dbVariants = list()

	def load(self, filePath = None):
		"""
			Prepare a hash table containing the keywords/phrases for search
		"""

		if filePath == None:
			filePath = self.dbPath

		fileIn = open(filePath, 'r')
		fileDB = fileIn.read()
		fileIn.close()

		for row in fileDB.split('\n'):
			if len(row) > 0 and row[6].lower() == 'i':
				row = row.split('\t')
				
				DICECode = int(row[0][2:5])
				variants = row[1]

				self.dbList.append((DICECode, variants))

	def getDB(self):
		return self.dbList

	def size(self):
		return len(self.dbList)

	def isSufficient(self, dice, term):
		if (dice, term) in self.dbList:
			return True
		return False

if __name__ == "__main__":

	timeOne = time.clock()
	baseDir = "M:\\DICE\\Trinity\\Extracts\\stpeters\\St Peters eMD Re-Tune Doc Extracr\\Extract1\\samples\\Sample_StPeters_Glenn-2000-deid\\cln\\"
	listDir = os.listdir(baseDir)

	ldb = LoadDB()
	ldb.load("L:\\DICE Documents\\JoshsAwesomeScript\\search-keywords-all.txt")
	print ldb.size()
	print ldb.isSufficient(6, "chf")
	print ldb.getDB()[0]

	print "DONE"