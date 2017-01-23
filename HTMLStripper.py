#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name: 	HTMLStripper.py
# Version: 	1.0.0
# Author: 	Glenn Abastillas
# Date: 	March 08, 2016
#
# Purpose: Allows the user to:
#           1.) Navigate a folder and its subfolders to load files intended for HTML mark-up removal
#			2.) Remove HTML mark-up with native HTMLParser class from each file encountered
#			3.) Save all documents with HTML mark-up removed into an 'output' folder in the save directory as the files
#
# To see the script run, go to the bottom of this page. 
#
# This class is directly inherited by:
#		- pyWeb/Web.py
#
# Updates:
# 1. [2016/03/10] - added documentation for all methods. Moved convert method out of the scripting pathway. Added HTMLStripperObject and LoopDirObject parameters to convert method.
# 2. [2016/03/11] - added clear_tabs() and is_blank() methods for use in strip_tags() for-loop.
# 3. [2016/03/11] - simplified generators to a single generator using methods added in update 2. Checked generator run time versus expanded for-loop (generator runs ~1.12x faster).
# - - - - - - - - - - - - -
"""	load files from a directory for HTML mark-up removal.

Using the built-in Python HTMLParser package, text data with HTML mark-up is fed into the HTMLParser class's feed method, which removes the mark-up. Post-processing is applied to remove extraneous empty lines and leading tabs.

This class searches the specified directory and its sub-directories entirely for all text files regardless of their HTML mark-up content and processes them. If a document has no HTML mark-up, then it simply passes through the post-processing stage.

To use this class as a script, simply change the variable (i.e., f1) pointing to the directory in the scripting section and run the script. A new 'output' folder will be created where the documents with HTML mark-up are found and the processed files are then saved at that location.
"""
from LoopDir 	import LoopDir

from HTMLParser import HTMLParser
import os

__author__ 		= "Glenn Abastillas"
__copyright__ 	= "Copyright (c) March 08, 2016"
__credits__ 	= "Glenn Abastillas"

__license__ 	= "Free"
__version__ 	= "1.0.0"
__maintainer__ 	= "Glenn Abastillas"
__email__ 		= "a5rjqzz@mmm.com"
__status__ 		= "Production"

class HTMLStripper(HTMLParser):

	def __init__(self):
		self.reset()
		self.fed = []

	def clear_tabs(self, string):
		""" used to remove extraneous tabs '\t' in a string
			@param	string - String text
			@return	String - string without extranous tabs
		"""
		# If there are '\t' characters in the string, replace them with blanks
		if '\t' in string:
			return string.replace('\t', '')
		else:
			return string

	def get_data(self):
		"""	return a string of the self.fed list
			@return	String - joined self.fed list
		"""
		return ''.join(self.fed)

	def handle_data(self, d):
		"""	save data with HTML mark-up removed to self.fed list
			@param	d - data output from the HTML mark-up removed process
		"""
		self.fed.append(d)

	def is_blank(self, string):
		"""	used for list comprehensions to see if an string is blank
			@param	string - String text
			@return	boolean - True if string is blank, else False
		"""
		return len(string.strip())==0

	def strip_tags(self, fileIn):
		"""	read in specified file in fileIn and remove HTML mark-up
			@param	fileIn - path to file for mark-up removal
			@return	String - of the post-processed file
		"""
		# Open and read in the specified file
		fileIn = open(fileIn, 'r')
		html = fileIn.read()
		fileIn.close()

		return self.strip(html)

	def strip(self, text):
		"""	remove HTML tags from an input text
			@param	text: text with HTML tags
			@return	String text without HTML tags
		"""
		
		# Send the file to the HTMLParser to extract all HTML tags, results use handle_data() to put output in self.fed
		self.feed(text)
		output = self.get_data().splitlines()
		output = (self.clear_tabs(L) for L in output if not self.is_blank(L))

		# Close HTML feed and clear self.fed in preparation for next document
		self.close()
		self.fed = []
		return '\n'.join(output)

def convert(directory, HTMLStripperObject, LoopDirObject):
	"""	loop through this directory and all subdirectories to apply the 
		HTMLStripper on.
		@param	HTMLStripperObject - instance of the HTMLStripper
		@param	LoopDirObject - instance of the LoopDir pointed to path
		@return None - end recursion process
	"""
	thisDirectory = os.listdir('.')

	hasFolder = bool(sum((True if os.path.isdir(f)  else False for f in thisDirectory)))
	hasFile   = bool(sum((True if os.path.isfile(f) else False for f in thisDirectory)))
	dirList   = (f for f in thisDirectory if os.path.isdir(f))

	print "hasFolder: {0} \thasFile: {1} \tCurrently Processing: {2}".format(hasFolder, hasFile,  os.getcwd())

	if hasFolder:

		# If this directory has folders, loop through them first until a file only directory is reached
		for D in dirList:

			# If the directory found is not an output folder, move into it and analyze again
			if D != "output":
				os.chdir(D)
				convert(D, HTMLStripperObject, LoopDirObject)
				os.chdir("..")

	elif hasFile:

		# If this directory has only files, apply the HTMLStripper.strip_tags method to each file
		tagless = LoopDirObject.apply(function=HTMLStripperObject.strip_tags)

		# If there is not already an output folder in this directory, create one
		if not os.path.isdir("output".format(directory)):
			os.mkdir("output".format(directory))

		# Base number for new file names
		i = 1000000

		# Write each of the files to the output folder in current directory
		for T in tagless:
			fout = open("output\\{0}.txt".format(i), 'w')
			fout.write(T)
			fout.close()
			i+=1

	return None
if __name__=="__main__":
    """ run as a script if this file is run as a stand-alone program

    	change variable f1 to point to the directory containing all the files
    	and folders that contain files with HTML mark-up needing removal
    """
    f1 = "M:\\DICE\\site - Parkland\\Extract1\\00_not_useful\\Anes\\Addendum Note"
    convert(f1, HTMLStripper(), LoopDir(f1))

