#!C:\Python27\python.exe
### -*- encoding: utf-8 -*-
### BatchReader.py
### Glenn Abastillas
### February 2, 2016
### Version 1.0
### BatchReader converts a multiple *.rtf files into *.txt and places them in a 'converted' folder.
### This class makes use of Rtf15Reader and PlaintextWriter from the pyth plugin downloaded here: https://pypi.python.org/pypi/pyth/0.5.4
### There are no methods or functions in this class. All processes run automatically upon instantiation.

from pyth.plugins.rtf15.reader 		import Rtf15Reader	
from pyth.plugins.plaintext.writer 	import PlaintextWriter
import os

class BatchReader(object):

	"""
			BatchReader converts multiple *.rtf files into *.txt files and places them into a 'converted' folder.
		This object does not contain any methods or functions. All process automatically execute upon instantiation.
	"""

	def __init__(self, filePath):

		# Create a list containing names for each *.rtf file located at the filePath location.
		rtfFiles = [f for f in os.listdir(filePath) if f[0] != '~' and '.rtf' in f]

		# Create a String indicating the path to the 'converted' folder.
		convertFolder = "{0}{1}".format(filePath, "converted\\")
		corruptFolder = "{0}{1}".format(filePath, "corrupted\\")

		# Check for existence of 'converted' folder. If it does not exist, a new 'converted' folder will be
		# created in the filePath specified. Otherwise, continue to the conversion process.
		if os.path.exists(convertFolder) == False:
			print "'converted' folder not found. Creating 'converted' folder ..."
			os.mkdir(convertFolder)
			print "'converted' folder created."
		else:
			print "'converted' folder found."

		try:
			# Iterate through the file names in rtfFiles to open in the *.rtf format for conversion in the
			# next step.
			print "Reading in *.rtf files."
			rtfFilesReadIn 		= (Rtf15Reader.read(open("{0}{1}".format(filePath,f), "rb")) for f in rtfFiles)

			# Iterate through the files read into the script in the *.rtf format and convert each file into a
			# plain text (i.e., *.txt) file for output.
			print "Converting *.rtf files to *.txt files."
			rtfFilesConverted 	= (PlaintextWriter.write(f) for f in rtfFilesReadIn)

			# Iterate through the converted files, create an associated file name and save into the 'converted'
			# folder created or found in prior steps. Keep track of the number of files for the final display.
			print "Writing *.txt files to {0}".format(convertFolder)
			
			count = 0

			for i, textFile in enumerate(rtfFilesConverted):
				outputFileName = "{0}{1}.txt".format(convertFolder, rtfFiles[i].split('.')[0])
				
				if os.path.exists(outputFileName) == False:
					output = ''.join(textFile.readlines())
					writeOut = open(outputFileName, 'w')
					writeOut.write(output)
					count += 1

		#except(IndexError, WrongFileType):
		finally:
			print "[Error Correction] Re-attempting conversion. Non-convertible files will be skipped and moved."
			### In the event of an error, loop through the files "manually" (i.e., not with a generator) and
			### file unconvertable files into a separate folder.

			### Iterate through the file names in rtfFiles to open in the *.rtf format for conversion in the
			### next step.
			corruptedFiles = []

			print "[Error Correction] Reading in *.rtf files."
			rtfFilesReadIn = []

			for f in rtfFiles:
				name = "{0}{1}".format(filePath, f)
				try:
					rtfFilesReadIn.append(Rtf15Reader.read(open(name, "rb")))
				finally:
					corruptedFiles.append(name)

			### Iterate through the files read into the script in the *.rtf format and convert each file into a
			### plain text (i.e., *.txt) file for output.
			print "[Error Correction] Converting *.rtf files to *.txt files."
			# -- rtfFilesConverted = (PlaintextWriter.write(f) for f in rtfFilesReadIn)
			rtfFilesConverted = []

			for f in rtfFilesReadIn:
				rtfFilesConverted.append(PlaintextWriter.write(f))

			### Iterate through the converted files, create an associated file name and save into the 'converted'
			### folder created or found in prior steps. Keep track of the number of files for the final display.
			print "[Error Correction] Writing *.txt files to {0}".format(convertFolder)
			
			count = 0

			for i, textFile in enumerate(rtfFilesConverted):
				outputFileName = "{0}{1}.txt".format(convertFolder, rtfFiles[i].split('.')[0])

				if os.path.exists(outputFileName) == False:
					output = ''.join(textFile.readlines())
					writeOut = open(outputFileName, 'w')
					writeOut.write(output)
					count += 1

			corruptCount = 0
			if len(corruptedFiles) > 0:
				### If there are any files that do not convert, then place them into a separate folder for manual
				### conversion. A new 'corrupted' folder will be created.

				### Create a "corrupted" folder for corruptedFiles
				if os.path.exists(corruptFolder) == False:
					os.mkdir(corruptFolder)

				### Place all the corrupted files in corruptedFiles into the corruptFolder
				writeOut = [open(cf, 'w') for cf in corruptedFiles]

				for i in xrange(len(corruptedFiles)):
					writeOut[i].write(corruptedFiles[i])
					writeOut[i].close()
					corruptCount += 1

			print "{0} corrupted files moved from {1} to {2}".format(corruptCount, filePath, corruptFolder)


		### Display completion message. If run from *.pyc, this process will allow the user to view the program's
		### output before closing out. Display number of files converted and location of saved converted files.
		print "{0} files converted from *.rtf to *.txt in {1}".format(str(count), convertFolder)
		# -- exit = input("Press any key to continue...")

		# Move *.rtf files into an 'rtf' folder after processing. Create 'rtf' folder if the folder does not already
		# exist.

		rtfFolder = "{0}{1}".format(filePath, "rtf\\")
		if os.path.exists(rtfFolder) == False:
			os.mkdir(rtfFolder)

		# Initialize count for moved files to 0.
		countMovedFiles = 0

		# Loop through the *.rtf files in this folder to be renamed (i.e., removed)
		for f in rtfFiles:
			oldPath = "{0}{1}".format(filePath, f)
			newPath = "{0}{1}".format(rtfFolder, f)
			os.rename(oldPath, newPath)		

		print "{0} files moved from {1} to {2}".format(str(countMovedFiles), filePath, rtfFolder)

if __name__ == "__main__":
	
	### *** YOU MAY CHANGE THIS VARIABLE ***
	### The variable filePath points to the folder location containing the *.rtf files to be converted
	### into plain text files (i.e., *.txt). 
	##filePath = "M:\\DICE\\site - Kettering\\Extract1\\samples\\Sample_Kettering_Glenn-1500-deid\\"

	baseFile = "M:\\DICE\\site - Kettering\\Extract1\\PHI\\raw\\20160202\\usable\\"

	#Folders with errors:filePath = ["CONSULTS\\consults\\", "PROG NOTES\\Progress Notes\\"]
	#All folders:filePath = ["DISCHARGE\\Discharge Summaries\\", "ED SUMMARY\\ED Notes\\", "ED SUMMARY\\ED Provider Notes\\", "H AND P\\History and Physical\\", "H AND P\\Interval History and Physical Note\\", "L AND D\\Labor and Delivery Note\\", "PROCEDURE\\Procedures\\", "PROG NOTES\\Progress Notes\\", "PROG NOTES\\Significant Event\\", "PROG NOTES\\Treatment Plan\\", "SURGICAL\\Brief Op Note\\", "SURGICAL\\Op Note\\"]
	#20160203 files: filePath = ["PROG NOTES\\Progress Notes\\", "PROG NOTES\\Significant Event\\", "PROG NOTES\\Treatment Plan\\", "SURGICAL\\Brief Op Note\\", "SURGICAL\\Op Note\\"]
	filePath = ["CONSULTS\\consults\\", "PROG NOTES\\Progress Notes\\"]

	### *** Script Starts Here *** 
	### Run the BatchReader script, passing the filePath variable into the class.
	##BatchReader(filePath)

	for f in filePath:
		BatchReader(baseFile+f)

