# Name: Scan.py
# Author: Glenn Abastillas
# Date: 8/21/2015
# Purpose: Allows the user to:
#           1.) Count the number of directories and subdirectories, etc., in a directory.
#           2.) Count the number of files in a directory and its subdirectories.
#
# To see the script run, go to the bottom of this page.
#
# This class is used in the following classes:
#
# - - - - - - - - - - - - -
import os

class Scan(object):

    def countFiles(self, directory = "", fileCount = 0):
        """
            countFiles() goes through all the folders in the directory indicated in 'directory' and counts the total number of files it contains.

            parameters
            directory:  parent/top-level node whose children are counted
            fileCount:  count of files prior to starting the method
                        default is initialized to 0

        """

        if os.path.isdir(directory):
            if self.hasOnlyFiles(directory) == True:

                return len(os.listdir(directory)) + fileCount

            elif self.hasOnlyFiles(directory) == False:

                folders = self.getFolderList(directory)

                os.chdir(directory)

                # CALL THESE JUST ONCE BEFORE LOOP(S)
                countFiles  = self.countFiles
                format      = str.format
                # - - - - - - - - - - - - - - - - - -

                for folder in folders:

                    folder      = directory + "\\" + folder
                    fileCount   = countFiles(folder, fileCount)

                    print format("Counting ~{0}", folder[len(directory):])

                os.chdir('..')

                fileCount = fileCount + len(self.getFileList(directory))

        return fileCount

    def hasOnlyFiles(self, directory = ""):
        if os.path.isdir(directory):

            isdir = os.path.isdir

            for item in os.listdir(directory):
                item = directory + "\\" + item

                if isdir(item) == True:

                    return False

        return True

    def getFolderList(self, directory = ""):
        """
            getFolderList() returns a list of all the folders in the directory specified.

            directory:  location whose folders the user wishes to get as a list.
        """
        return [item for item in os.listdir(directory) if os.path.isdir(directory + "\\" + item) == True]

    def getFileList(self, directory = ""):
        """
            getFileList() returns a list of all the file in the directory specified.

            directory:  location whose file the user wishes to get as a list.
        """
        return [item for item in os.listdir(directory) if os.path.isfile(directory + "\\" + item) == True]

if __name__=="__main__":
    import InfoText, time

    i = InfoText.InfoText()
    t = Scan()
    
    print i.scan()
    print i.separator()
    print i.location()
    print i.separator()

    print "M:\\DICE\\ ..."
    user_input = "M:\\DICE\\" + raw_input("Enter hospital directory whose files you would like to count:\n\n>>\t")
    
    startTime = time.clock()
    
    print "\nCounting files in {0}...\n".format(user_input)
    print t.countFiles(directory = user_input)
    print "\nCounting took {0} seconds".format(time.clock()-startTime)

    exit = raw_input("File count complete. Press any key to exit ...\n")

