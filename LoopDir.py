#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:	 LoopDir.py
# Version:  1.0.2
# Author:   Glenn Abastillas
# Date:	 September 16, 2015
#
# Purpose: Allows the user to:
#		   1.) Open a specified directory.
#		   2.) Loop through listings applying a function on files and folders if applicable.
# 
# To see the script run, go to the bottom of the page.
#
# This class is directly inherited by the following classes: 
#   - ..\Analyze.py
#   - Scrambler.py
#
# Updates:
# 1. [2015/12/07] - optimized for loops in apply() and string concatenation in save() methods. Version changed from 1.0.0 to 1.0.1.
# 2. [2016/02/29] - changed wording of notes in line 14 from '... class is used in the following ...' to '... class is directly inherited by the following ...'. Print '[ERROR]' statement added to line 112 in apply() if no function specified. Version changed from 1.0.1 to 1.0.2.
# 3. [2016/03/09] - changed unit variable to equal the result of a ternary-like operation (e.g., len/10 if len/10 else 1).
# 4. [2016/03/09] - added print statement at the end of the loop in apply method to indicate completion (i.e., <__name__> Complete).
# 5. [2016/03/09] - removed the returnResults parameter from the apply method. Refactored logic so the if statemtents looks at printOut instead. If printOut==True then print, else return results.
# 6. [2016/03/11] - added a spinning cursor animation to the progress bar at the end of the apply() method.
# 7. [2016/03/14] - added recursive loop method.
# 8. [2016/03/14] - commented out if-logic for directory==None in apply method.
# - - - - - - - - - - - - -
""" creates an object of a directory indicated during construction or through setDir().

An object of the specified directory is made by setting the directory at object construction or through the setDir method. Once the object is created, the apply method allows for a function to work on all the files or folders in the specified directory.

In the apply() method, cycle through directory as indicated by "directory". If there is a function assigned, each item will be pass through that function.

Functions must be passed through as: Class().Function
				
Omit "()" after the Function (cf., Class().Function()).
Here's an example of this format:
	Animals().Eat*
					
	as in Directory.apply(Animals().Eat)

If this class has been instantiated to 'a':
	a = Animals() --> a.Eat*

	*as in Directory.apply(a.Eat)
"""
import os, sys
import time

__author__ 		= "Glenn Abastillas"
__copyright__ 	= "Copyright (c) September, 16, 2015"
__credits__ 	= "Glenn Abastillas"

__license__ 	= "Free"
__version__ 	= "1.0.2"
__maintainer__ 	= "Glenn Abastillas"
__email__ 		= "a5rjqzz@mmm.com"
__status__ 		= "Production"

class LoopDir(object):

	def __init__(self, directory = None):
		"""	constructor for instance.
			@param	directory - folder location
		"""
		self.directory	  	= directory	# Path to directory
		self.directorySet   = False		# If directory is set --> True
		self.directoryList  = list()	# List of files and folders in directory
		self.iterationIndex = 0			# For iterable

		# If directory is specified, set directory
		if directory is not None:
			self.setDir(directory)

	def __getitem__(self, key):
		"""	returns item in list as indicated by key number. 
			@param key	index of the item in the list.
			@return directoryList item at specified index
		"""
		try:
			return self.directoryList[key]
		except(IndexError):
			return IndexError("Directory item does not exist.")
		
	def __iter__(self):
		"""	returns class as part of loop in conjunction with next().
			@return	return self
		"""
		return self

	def next(self):
		"""	advance to the next item in a list. 
			Stops iteration when it reaches the end.
			@return	directoryList - item at specified index
		"""
		try:
			self.iterationIndex += 1
			return self.directoryList[self.iterationIndex - 1]
		except(IndexError):
			self.iterationIndex = 0
			raise StopIteration
		
	def setDir(self, directory = '.'):
		"""	open directory as indicated by "directory". If nothing indicated, 
			stays in current working directory.
			@param	directory - location of files and folders to loop through
		"""
		
		if self.directory is None:
			self.directory = directory
			
		os.chdir(self.directory)
		self.directoryList = os.listdir(self.directory)
		self.directorySet  = True

	def apply(self, function = None, directory = '.', actOnFiles = True, printOut = False, appendToList = True):
		"""	apply specified function to files or folders in specified directory.
			@param functions  - function to be applied
			@param actOnFiles - act on file items if True
			@param directory  - location of folders and files to loop through
			@param printOut   - print to screen or return results
		"""
		
		if function is None:
			return ValueError("No function specified. No processes started. Nothing returned.")

		directoryList = list()#os.listdir(directory)
		outputResults = list()

		# If False, apply function to folders. Otherwise, apply function to files.
		if actOnFiles == False:
			isGood = os.path.isdir
			
		else:
			isGood = os.path.isfile

		for item in os.listdir(directory):
			if isGood(item):
				directoryList.append(item)

		directoryList = sorted(directoryList, key=lambda item: item.split('.')[-1], reverse=True)


		# CALL THESE JUST ONCE BEFORE LOOP(S)
		append = outputResults.append
		#isGood = os.path.isfile
		index  = directoryList.index
		unit   = len(directoryList)/10 if len(directoryList) >= 10 else 1


		# Loop through the 
		for item in directoryList:
			if isGood(item):
				if appendToList:
					append(function(item))
				else:
					function(item)

			i = index(item)

			if i%unit == 0:
				for C in ('\b\\', '\b|', '\b/', '\b\\', '\b|', '\b/', '\b\\', '\b|', '\b/', '\b\\', '\b|', '\b/', '\b.'):
					sys.stdout.write(C)
					time.sleep(0.05)
					sys.stdout.flush()
				print '|',

		print "| ({0}) Complete".format(function.__name__)

		self.directoryList = directoryList
		self.directorySet  = True
			
		if printOut==False:
			return outputResults
		print outputResults

	def recursiveApply(self, function=None, directory='.', actOnFiles = True, printOut = False, appendToList = True):
		"""	apply specified function on all files in the directory specified
			and all of its subdirectories if present.
			@param functions  - function to be applied
			@param actOnFiles - act on file items if True
			@param directory  - location of folders and files to loop through
			@param printOut   - print to screen or return results
		"""

		directoryList = os.listdir(directory)
		directoryList = sorted(directoryList, key=lambda item: item.split('.')[-1], reverse=True)

		# CALL THESE JUST ONCE BEFORE LOOP(S)
		#append = outputResults.append
		isGood = os.path.isdir#file
		#index  = directoryList.index
		#unit   = len(directoryList)/10 if len(directoryList) >= 10 else 1
		# - - - - - - - - - - - - - - - - - -

		#for D in directoryList:
		#	if isGood(D):
		#		pass

	def saveDir(self, saveToFileName = "Directory_Class_Output", saveToFileDir = "Directory_Class_Output", fileType = "txt"):
		"""	save the directory listing as a text file. 
			Either apply() or setDir() must be run to save.
			
			@param saveToFileName Desired name of output file
			@param saveToFileDir  Desired name of output directory
			@param fileType	  Desired file type of output file
		"""
		if self.directorySet:
			try:
				os.mkdir(saveToFileDir)
			except(WindowsError):
				pass

			os.chdir(saveToFileDir)
			
			saveToFileTime = time.strftime("_%Y%m%d_%H%M%S")
			saveToFileName = "{0}{1}.{2}".format(saveToFileName, saveToFileTime, fileType)
			saveToFileData = str(self.directoryList)

			saveToFile = open(saveToFileName, 'w')
			saveToFile.write(saveToFileData)
			saveToFile.close()
			
			os.chdir("..")
		else:
			print "No directory set. Set directory through setDir() or apply() functions."
			
if __name__ == "__main__":
	"""
	import Document
	class test():
	def test(self, string):
		t = len(string) + 1
		fo = open(str(t) + "_testTest.txt", "w")
		fo.write(str(t) + " this is a test " + str(t))
		fo.close()
		print t
		
	#"""

	class Lens(object):

		def getLen(self, inputs):
			return inputs, len(inputs)

	l = Lens()
	#dc = Document.Document()
	d = LoopDir("M:\DICE\Hendrick\Extract1\samples\Sample_Hendrick_Glenn-1500-deid\cln")
	d.apply(l.getLen)
	d.saveDir()
	
	#d.apply(test().test)
	#d.saveDir()
	#d.apply()
	#print d[-1]
