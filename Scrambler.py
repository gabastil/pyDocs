#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Scrambler.py
# Version:  1.0.0
# Author:   Glenn Abastillas
# Date:     September 16, 2015
#
# Purpose: Allows the user to:
#           1.) Perform the tasks of "PrepFiles5.pl" in "L:\DICE Documents\Scripts\prepfiles\P"
# 
# To see the script run, go to the bottom of the page.
#
# This class is not directly inherited by any other class.
#
# Updates:
# 1. [2015/10/01] DEPRECATED. Class discontinued.
#
# - - - - - - - - - - - - -

from LoopDir import LoopDir
import os, sys, time

class Scrambler(LoopDir):

        """
                This class performs the tasks of the PERL script "PrepFiles5.pl" in the directory 
                "L:\DICE Documents\Scripts\prepfiles\P". There are six stages:

                1. Rename files
                2. Sample files
                3. Generate new names
                4. Display time statistics

        """

        def __init__(self,      baseDrive       = "M:\\DICE\\", 
                                baseDir         = None,                 
                                docInfoFileName = "docinfo.txt", 
                                startingNumber  = 2000000, 
                                samples         = None, 
                                logFile         = "FilePrepLog.txt"):
        
                self.baseDrive       = baseDrive
                self.baseDir         = baseDir

                self.docInfoFileName = docInfoFileName

                self.logFile         = logFile
                
                self.startingNumber  = startingNumber
                self.samples         = samples

                self.fileNameList    = []

        def renameFiles(self):
                self.isBaseDirNone()
                os.chdir(self.baseDir)
                listOfFilesInDir = os.listdir('.')
                pass

        def sampleFiles(self):
                pass

        def generateNewNames(self):
                pass

        def numThenChar(self):
                pass

        def getFilesFromDir(self):
                pass

        #=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*       
        #               THESE METHODS SERVICE THIS CLASS ONLY (BELOW)
        #=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*

        def isBaseDirNone(self):
                if self.baseDir == None:
                        raise ValueError("No base directory indicated.")
                        return True
                else:
                        self.baseDir = self.baseDrive + self.baseDir
                        return False

        def startTime(self):
                self.time = time.clock()

        def stopTime(self):
                return time.clock() - self.time

        def rounds(self):
                pass

        #=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*       
        #               THESE METHODS SERVICE THIS CLASS ONLY (ABOVE)
        #=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*

if __name__ == "__main__":
        import os
        baseDir = "Halifax\\Extract1\\txt"

        #s = Scrambler(baseDir = "Halifax")

        s = Scrambler(baseDir = baseDir)

        s.startTime()
        print os.listdir('.')
        print s.stopTime()
        s.renameFiles()
        pass
