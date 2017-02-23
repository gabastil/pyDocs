#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Spreadsheet.py
# Version:  1.0.1 (January 26, 2017)
# Author:   Glenn Abastillas
# Date:     August 21, 2015
#
# Purpose: Allows the user to:
#           1.) Load a spreadsheet into memory.
#           2.) Transpose columns and rows.
#           3.) Find a search term and return the column and row it is
#               located in.
#
# This class does not have scripting code in place.
#
# This class is directly inherited by the following classes:
#       - SpreadsheetPlus.py
#
# Updates:
# 1. [2015/12/04] - added: method open().
# 2. [2016/02/09] - in load() method, removed lines 125 - 128 including
#                   else-statement. Moved file_in from inside nested
#                   else-statement.
# 3. [2016/02/29] - changed wording of notes in line 17 from '... class
#                   is used in the following ...' to '... class is directly
#                   inherited by the following ...'.
# 4. [2016/03/07] - removed unused method - def determineGroupByRange(self,
#                   columnToBeDetermined = 0, rangeToBeDetermined = 0)
# 5. [2016/04/18] - added: addColumn(), newColumn() and fillColumn() methods
# 5. [2016/04/21] - added: addColumn(), newColumn() and fillColumn() methods
# 6. [2016/12/10] - added: getCell(), getRowCount() and getColumnCount()
#                   methods
# 7. [2016/12/12] - added blocks of code to allow for generator input for
#                   setData() method
# 8. [2017/01/20] - reformatted entire class according to PEP guidelines.
#                   Class notes to be completed at a later time.
# 8a.[2017/01/20] - added __repr__() class.
# 9. [2017/01/24] - collapsed setters and getters for cell(), row(), column()
# 9a.[2017/01/24] - changed fillColumn() to fill(), initialize() to load()
# 9b.[2017/01/24] - collapsed setters and getters for savePath() and filePath()
# 9c.[2017/01/24] - updated load() method
# 10.[2017/01/25] - added is_[state]() methods for loaded, initialized, and
#                   transposed states.
# 10.[2017/01/26] - completed Documentation for the Spreadsheet object. Updated
#                   notes for several methods. Fixed the following bugs:
#
#                   getColumnName() - method was incomplete. Method now gets
#                            the column name correctly.
#
#                   load() - Did not assign headers causing getHeaders()
#                            to retrieve the first non-header row. Assigns
#                            correct headers now.
#                            self.initialized now works for all load cases.
# 11.[2017/02/19] - Added getState() method and revertTranspose() methods.
# - - - - - - - - - - - - -

__author__ = "Glenn Abastillas"
__copyright__ = "Copyright (c) August 21, 2015"
__credits__ = "Glenn Abastillas"

__license__ = "Free"
__version__ = "1.0.0"
__maintainer__ = "Glenn Abastillas"


def preserve_transpose(function):
    """
    Allows for preserving the original transposition of the
    functions in Spreadsheet

    Args:
        function (function): wrapped function
        args (tuple): arguments to be used by function
        kwargs (dict): keyword arguments to be used by function

    Returns:
        whatever the wrapped function returns
    """
    def preserve(*args, **kwargs):
        # print args[0].getState()
        state = args[0].getState()
        results = function(*args)
        args[0].revertTranspose(state)
        return results

    return preserve


class Spreadsheet(object):

    """
        Spreadsheet class creates a table like data structure for other classes
        to load spreadsheets specified character separated values to be manipu-
        lated, i.e., searched, rows/columns edited, and saved.

        Data input can be manipulated (e.g., tranposed, data addition, and rem-
        oval) and saved to a specified output file. Data delimiters are specif-
        ied as required with the default delimiter being tab-delimited '\t'.

        User Accesible Methods:

            Methods for ROWS:

                row(row (int))
                    --> Gets row at specified index.

                addToRow(row (int), data (str, int))
                    --> Adds value to last empty cell in row.

                getRowCount()
                    --> Gets a count of rows in spreadsheet.

                getRowIndex(row (list))
                    --> Gets index of specified row in spreadsheet.

                removeRow(row (int))
                    --> Removes specified row.

                toRows()
                    --> Transposes spreadsheet so each list item is a row.

            Methods for COLUMNS:

                addToColumn(column (int), data (str, int))
                    --> Adds value to last empty cell in column.

                fill(column (int), fillWith (str, int))
                    --> Fills entire specified column with a single value
                        leaving the header intact.

                column(column (int), data (list),
                       number (bool), header (bool))
                    --> Gets column as a list using column name or index.
                    --> Gets column with header if header is True.
                    --> Gets column with numbers converted to number type if
                        number is True.
                    --> Replaces column at specified index if data provided.

                getColumnCount()
                    --> Gets a count of columns in spreadsheet.

                getColumnIndex(column (str, int))
                    --> Gets index of column if name of column is known.
                    --> Integer indices will return as themselves.

                getColumnName(column (int))
                    --> Gets name of column if index of column is known.

                newColumn(name (str), fillWith (str, int))
                    --> Creates a new column.

                rename(name (str), column (int))
                    --> Renames specified column.

                removeColumn(column (int))
                    --> Removes specified column.

                toColumns()
                    --> Transposes spreadsheet so each list item is a column.

            Methods for CELLS:

                cell(row (int), column (int), data (str, int))
                    --> Gets value at specified row and column.
                    --> Sets value at specified row and column if data present.

            Methods for DATA:
                getHeaders()
                    --> Gets the headers in this spreadsheet.

                getSpreadsheet()
                    --> Gets the spreadsheet data.

                setData(data (Spreadsheet, list))
                    --> Replaces the current spreadsheet with new data.

                transpose()
                    --> Transposes spreadsheet so rows --> columns and
                        vice-versa.

                toString(fileToString (str))
                    --> Returns a formatted string of the spreadsheet or data
                        specified.

                sort(column (int), reverse (bool), hasTitle (bool))
                    --> Sorts entire spreadsheet by column. Reverse sort done
                        if reverse is True. Title is also sorted if hasTitle
                        is True.

            Methods for FILE PATHS:
                filePath(filePath (str), delimiter (str))
                    --> Sets spreadsheet's file path and loads the spreadsheet.

                savePath(savePath (str))
                    --> Sets where the save() method will save the spreadsheet

                prepareForSave(spreadsheet (list), delimiter (str))
                    --> Formats specified spreadsheet for saving.

                open(filePath)
                    --> Opens specified file and returns a list.

                load(filePath (str), delimiter (str))
                    --> Opens specified file and sets state for Spreadsheet.

                refresh()
                    --> Adds addition cell padding if rows or columns are of
                        unequal lengths.

                reset()
                    --> Resets Spreadsheet back to its initial state with no
                        spreadsheet loaded.
    """

    COLUMN_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    def __init__(self, filePath=None, savePath=None,
                 delimiter="\t", columns=["columnName"]):
        """
            Initializes an instance of this class

            Attributes:
                filePath (str): path of the spreadsheet file to be loaded
                savePath (str): write to this location
                delimiter (str): spreadsheet delimiter
                columns (array): list of column names for a blank spreadsheet
        """

        # list containing spreadsheet
        self.spreadsheet = list()

        # location of the spreadsheet
        self.filePath = filePath

        # save location for spreadsheet
        self.savePath = savePath

        # spreadsheet loaded?
        self.loaded = False

        # spreadsheet initialized?
        self.initialized = False

        # spreadsheet as rows (False) or columns (True)
        self.transposed = False

        # Used for next() function
        self.iter_index = 0

        # Load filePath if one is specified
        if filePath is not None:
            self.load(filePath, delimiter)
        else:
            self.spreadsheet.extend([columns])

    def __getitem__(self, key):
        """
            Enables list[n] syntax

            Attributes:
                key (int): index number of item.
        """
        return self.spreadsheet[key]

    def __setitem__(self, i, item):
        """
            Enable list[n] = item assignment

            Attributes:
                i (int): index of item to replace
                item (obj): new item to assign
        """
        self.spreadsheet[i] = item

    def __iter__(self):
        """
            Enable iteration of this class
        """
        return self

    def __len__(self):
        """
            Returns number of items in spreadsheet
        """
        return len(self.spreadsheet)

    def __repr__(self):
        """
            Formats the string representation of this object
        """
        return self.toString()

    def next(self):
        """
            Get next item in iteration
        """
        try:
            self.iter_index += 1
            return self.spreadsheet[self.iter_index - 1]
        except(IndexError):
            self.iter_index = 0
            raise StopIteration

    @preserve_transpose
    def addToColumn(self, column=-1, data=None):
        """
            Fill a column with text

            Attributes:
                column (int): column to fill (index or string)
                data (str, int): data to insert
        """
        # -- state = self.getState()
        index = self.getColumnIndex(column)

        self.toColumns()
        lastEmptyCellIndex = self.getLastEmptyCell(column)
        columnForAppending = self.spreadsheet[index]

        if lastEmptyCellIndex == 0:
            columnForAppending.append(data)
        else:
            columnForAppending[lastEmptyCellIndex] = data
        # -- self.revertTranspose(state)

    @preserve_transpose
    def addToRow(self, row=-1, data=None):
        """
            Append or extend data to a row

            Attributes:
                row (int): index of row to add data to
                data (list): data to insert
        """
        # -- state = self.getState()
        self.toRows()
        if isinstance(data, list):
            self.spreadsheet[row].extend(data)
        else:
            self.spreadsheet[row].append(data)
        # -- self.revertTranspose(state)

    @preserve_transpose
    def cell(self, row, column, data=None):
        """
            Get or set value at specified cell

            Attributes:
                row (int): index of cell row
                column (int): index of cell column
                data (str): data to fill cell
        """
        if data is None:
            return self.getRow(row)[column]
        else:
            # -- state = self.getState()
            self.toRows()
            self.spreadsheet[row][column] = data
            # -- self.revertTranspose(state)

    def column(self, column=0, data=None, number=False, header=False):
        """
            Get column at specified index, add column if column is a list
            or replace column

            Attributes:
                column (int): index for desired column
                column (list): column to add
                data (list): new data to replace column at specified index
                number (bool): convert list elements to numbers if True
                header (bool): include header in list returned if True

            Returns:
                list: column elements
        """

        # Transpose spreadsheet to edit columns
        state = self.getState()
        self.toColumns()

        # If column parameter is an integer, retrieve column at index
        if isinstance(column, int) or isinstance(column, str):

            # Get the index of specified column - method used in case input
            # is string
            index = self.getColumnIndex(column)

            # Get base column
            column = self.spreadsheet[index]

            # Separate base column into head and body
            column_head = [column[0]]
            column_body = column[1:]

            # If there is data to set column, replace column with new data
            if isinstance(data, list):

                # If header is True, replace old header
                if header:
                    new_column = data

                # If header is False, keep old header
                else:
                    new_column = column_head + data

                self.spreadsheet[index] = new_column
                self.refresh()

            else:

                # Convert cell values to numbers
                if number is True:
                    body = [float(row) for row in column_body if row != '']

                    column = column_head + body

                    # Return column with converted values without header
                    if not header:
                        return body

                # Return the column without the header
                if not header:
                    return column_body

                return column

        # If column parameter is a list, add column to self.spreadsheet
        elif isinstance(column, list):

            self.spreadsheet.append(column)
            self.refresh()

        # If newColumn is none, raise an error
        elif column is None:
            raise ValueError("Please specify a new column list" +
                             "to add to the spreadsheet. E.g.," +
                             "Spreadsheet.addColumn(columnAsList)")

        self.revertTranspose(state)

    def filePath(self, filePath=None, delimiter="\t"):
        """
            Sets or gets the file path

            Attributes:
                filePath (str): path to load for this Spreadsheet
                delimiter (str): delimiter of document at filePath

            Returns:
                str: file path of this Spreadsheet
        """

        if filePath is None:
            return self.filePath
        else:
            self.load(filePath, delimiter)

    def fill(self, column=-1, fillWith=" ", skipTitle=True, cellList=None):
        """
            Fills a column with text

            Attributes:
                column (int): column to fill (index or column name)
                fillWith (str): datas of new column
                cellList (array): list of tuples (cell, offset) to
                                  insert into the fillWith formula
                skipTitle (bool): start filling on the second row
        """
        # Transpose spreadsheet to edit column
        state = self.getState()
        self.toColumns()

        initialIndex = 0
        newColumn = list()
        append = newColumn.append

        # if column variable is a string (i.e., column name),
        # get column with the same name's index
        if isinstance(column, str):
            column = self.getColumnIndex(column)

        # if skipTitle is True, start loop at index 1 not 0
        if skipTitle is True:
            append(self.spreadsheet[column][0])
            initialIndex = 1

        # if there is a a cellList, there are values to replace
        if cellList is not None:

            # items in cellList re tuples
            # the first part is the reference (e.g., "$A1" --> "$A{0}")
            # the second part is the row displacement, if any, (e.g., "-1")
            # for example,  [("$A{0}", "1"), ("$C${0}", "-1")] becomes:
            #               ["$A2", "$C$0"] at i-index == 1 in the loop

            cellListLength = len(cellList)
            columnForLoop = len(self.spreadsheet[column])

            # loop through the rows in the spreadsheet to fill data
            for i in xrange(initialIndex, columnForLoop):

                fillWithToReplace = fillWith

                # loop through the list of cells to insert into the
                # fillWith formula
                for j in xrange(cellListLength):

                    cellListValue = cellList[j][1]
                    # fillWithToReplace = cellList[j][0]

                    # if the cellList marker (i.e., {0}) to be replaced is a
                    # column the reference value has to be a letter

                    # -- THIS NEEDS TO BE FIXED -- #
                    if isinstance(cellListValue, str):
                        # -- referenceValue = self.COLUMN_ALPHA[column]
                        referenceValue = str(self.getColumnIndex(column,
                                                                 alphaIn=True))

                    # else, the reference value is a number
                    else:
                        referenceValue = str(i + cellListValue)

                    # reference cell to insert into the fillWith data
                    reference = cellList[j][0].replace("{0}", referenceValue)

                    # marker indicates where the reference should go
                    # (e.g., "{0} goes here and then {1}")
                    # -- NOT USED -- marker = "{}{}{}".format("{", str(j), "}")

                    # replace the marker in the fillWith string with the
                    # reference cell, e.g., "{0} goes here and then {1}";
                    # marker = "{0}"; reference = "$A$1"
                    # "$A$1 goes here and then {1}"
                    # fillWithToReplace = fillWithToReplace.replace(marker,
                    #                                               reference)
                    fillWithToReplace = reference
                    print fillWithToReplace

                append(fillWithToReplace)

        # if there is no cellList, just fill all cells with the same data
        else:

            columnForLoop = self.spreadsheet[column][initialIndex:]

            # loop through the rows in the spreadsheet to fill data
            for row in columnForLoop:
                append(fillWith)

        # print newColumn
        self.spreadsheet[column] = newColumn
        self.revertTranspose(state)

    @preserve_transpose
    def getLastEmptyCell(self, column=-1):
        """
        Return last empty cell for specified column

        Args:
            column (str, int): column to find last empty cell for

        Returns:
            int of last empty cell with spreadsheet set to columns
        """
        # -- state = self.getState()
        index = self.getColumnIndex(column)

        columnToCheck = self.column(index)

        lastEmptyCell = 0

        for i, row in enumerate(columnToCheck):
            if row == "":
                lastEmptyCell = -(i + 1)
            else:
                break

        # -- self.revertTranspose(state)
        return lastEmptyCell

    @preserve_transpose
    def getColumnCount(self):
        """
            Return number of columns in spreadsheet data
        """
        # -- state = self.getState()
        self.toColumns()
        columnCount = len(self.spreadsheet)
        # -- self.revertTranspose(state)
        return columnCount

    def getColumnIndex(self, column, alphaIn=False, alphaOut=False):
        """
            Get the numerical index for columns as per Excel coordinates

            Attributes:
                column (str, int): column character (e.g., 'A') or column name
                alphaIn (bool): indicate whether the string refers to excel
                                column headers (e.g., AA, AB, etc.)
                alphaOut (bool): indicates whether the output should be in
                                 excel column header format (e.g., AA)

            Returns:
                int: index of column specified
        """

        # If 'column' is a number type, return the integer of that number
        if isinstance(column, int) or isinstance(column, float):

            columnIndex = int(column)

            if alphaOut:

                columnIndex = ""

                for letter in str(column):
                    columnIndex += self.COLUMN_ALPHA[int(letter) - 1]

            return columnIndex

        # If 'column' is a single character, return it's index
        if alphaIn:
            column = column.upper()
            len_alpha = len(column)
            columnIndex = 0

            for i, letter in enumerate(column):
                letter_index = self.COLUMN_ALPHA.index(letter) + 1
                exponent = len_alpha - (i + 1)
                columnIndex += letter_index * (26 ** exponent)

            return columnIndex

        # If 'column' is a string, return it's index
        else:
            state = self.getState()
            self.toRows()
            columnIndex = self.spreadsheet[0].index(column)
            self.revertTranspose(state)
            return columnIndex

    @preserve_transpose
    def getColumnName(self, column):
        """
            Get the name of specified column

            Attributes:
                column (int): index of column

            Returns:
                str: name of column
        """
        # -- state = self.getState()
        index = self.getColumnIndex(column=column)
        self.toColumns()
        columnName = self.spreadsheet[index][0]
        # -- self.revertTranspose(state)
        return columnName

    @preserve_transpose
    def getHeaders(self):
        """
            Return headers for columns in the spreadsheet
        """
        # -- state = self.getState()
        self.toColumns()
        headers = [column[0] for column in self.spreadsheet]
        # -- self.revertTranspose(state)
        return headers

    @preserve_transpose
    def getRowCount(self):
        """
            Returns number of rows in spreadsheet data
        """
        # -- state = self.getState()
        self.toRows()
        rowCount = len(self.spreadsheet) - 1
        # -- self.revertTranspose(state)
        return rowCount

    @preserve_transpose
    def getRowIndex(self, data):
        """
            Get the index of a specified row

            Attributes:
                data (str):    cell data whose row to get

            Returns:
                int: index of row
                None: if data does not match row[0]
        """
        # -- state = self.getState()
        self.toRows()

        for row in self.spreadsheet:
            if data in row[0]:
                return self.spreadsheet.index(row)
        # -- self.revertTranspose(state)
        return None

    def getSpreadsheet(self):
        """
            Get this Spreadsheet
            @return List of rows and columns with data
        """
        return self.spreadsheet

    def getState(self):
        """
        Return state of transposition

        Returns:
            bool: True if transposed, False if not
        """
        return self.transposed

    def isInitialized(self):
        """
            Returns state of self.initialized
        """
        return self.initialized

    def isLoaded(self):
        """
            Returns state of self.loaded
        """
        return self.loaded

    def isTransposed(self):
        """
            Returns state of self.transposed
        """
        return self.transposed

    def load(self, filePath=None, delimiter="\t"):
        """
            Open the file and parse out rows and columns

            Attributes:
                filePath (str): spreadsheet file to load into memory
                delimiter (str): delimiter of document at filePath

            Raises:
                ValueError: if filePath is not specified
        """
        if filePath is None:
            raise ValueError("Please enter a file path for this method's" +
                             " filePath parameter")

        openedFilePath = self.open(filePath).splitlines()

        for line in openedFilePath:
            self.spreadsheet.append(line.split(delimiter))

        self.filePath = filePath
        self.loaded = True
        self.toRows()

        if not self.initialized and self.spreadsheet[0] == "columnName":
            del(self.spreadsheet[0])

        self.initialized = True

    def newColumn(self, name=" ", fillWith=" "):
        """
            Adds a new (empty) column to the spreadsheet

            Attributes:
                name (str): name of column
                fillWith (str): datas of new column
        """
        # Transpose spreadsheet to edit column
        self.toColumns()

        length = len(self.spreadsheet[0])

        newColumn = [fillWith for row in xrange(length)]
        newColumn[0] = name

        self.spreadsheet.append(newColumn)
        self.refresh()

    def open(self, filePath):
        """Opens an indicated text file for processing
            @param  filePath: path of file to load
            @return String of opened text file
        """
        with open(filePath, 'rUb', 2) as fileIn1:
            fileIn2 = fileIn1.read()
            fileIn1.close()

        return fileIn2

    def prepareForSave(self, spreadsheet=None, delimiter="\t"):
        """
            Prepares the spreadsheet for saving

            Attributes:
                spreadsheet (array): list of rows/columns to prepare

            Returns:
                str: String of spreadsheet in normal form (e.g. not transposed)
        """
        # Use this instance's spreadsheet if none specified
        if spreadsheet is None:
            self.toRows()
            spreadsheet = self.spreadsheet

        rowList = list()
        append = rowList.append

        for row in spreadsheet:
            row = [str(item) for item in row]
            append(delimiter.join(row))

        savedata = "\n".join(rowList)

        return savedata

    def refresh(self):
        """
            Make sure all columns/rows are the same length
        """
        self.transpose()
        self.transpose()

        if len(self) > 0:
            self.initialized = True

    def rename(self, name=None, index=-1):
        """
            Rename a specified column

            Attributes:
                name (str): name for the column
                index (int): column index
        """
        # Transpose spreadsheet to edit column
        self.toColumns()
        self.spreadsheet[index][0] = name

    def removeColumn(self, column=-1):
        """
            Removes a column in self.spreadsheet

            Attributes:
                column (int, str): index or string indicating column to remove
        """
        self.toColumns()

        # if column variable is a string (i.e., column name),
        # get column with the same name's index
        if isinstance(column, str):
            column = self.getColumnIndex(column)

        del self.spreadsheet[column]

    @preserve_transpose
    def removeRow(self, row=-1):
        """
            Removes a row in self.spreadsheet

            Attributes:
                row (int): index of row to remove
        """
        state = self.getState()
        self.toRows()

        del self.spreadsheet[row]
        self.revertTranspose(state)

    def reset(self):
        """
            Reset all data in this class
        """

        self.spreadsheet = list()  # list containing spreadsheet

        self.filePath = None  # location of the spreadsheet
        self.savePath = None  # location of the spreadsheet

        self.loaded = False  # spreadsheet loaded?
        self.initialized = False  # spreadsheet initialized?
        self.transposed = False  # checks if rows (=f) or columns (=t)

        self.iter_index = 0  # Used for next() function

    @preserve_transpose
    def row(self, row=0):
        """
            Get row at specified index or add row if row is a list

            Attributes:
                row (int): index for desired row
                row (list): row to add

            Returns:
                list: row elements
        """
        state = self.getState()
        self.toRows()

        # If row is an integer, get row at that index
        if isinstance(row, int) or isinstance(row, float):
            row = int(row)
            return self.spreadsheet[row]

        elif isinstance(row, list):

            # If row is none, raise an error
            if row is None:
                raise ValueError("Please specify a new row list" +
                                 "to add to the spreadsheet. E.g.," +
                                 " Spreadsheet.row(rowAsList)")

            # Transpose spreadsheet to edit row
            self.spreadsheet.append(row)
            self.refresh()

        self.revertTranspose(state)

    def save(self, savePath=None, savedata=None,
             saveType='w', delimiter="\t"):
        """
        Writes data out to a file

        Args:
            savePath (str): name of the file to be saved
            savedata (str): list of rows/columns to be saved
            saveType (str): indicate overwrite ('w') or append ('a')
            delimiter (str): type of delimiter to use for output
        """

        if savePath is None:
            savePath = self.savePath

        if savedata is None:
            savedata = self.prepareForSave(delimiter=delimiter)
        else:
            savedata = self.prepareForSave(spreadsheet=savedata,
                                           delimiter=delimiter)

        with open(savePath, saveType) as fout:
            fout.write(savedata)

    def setData(self, data, asRows=True):
        """
            Set spreadsheet to new data

            Attributes:
                data (Spreadsheet, list): user specified spreadsheet data
        """
        list_type = isinstance(data, list)
        spreadsheet_type = isinstance(data, Spreadsheet)

        # if data is neither a list or Spreadsheet, raise TypeError()
        if not list_type and not spreadsheet_type:
            raise TypeError("Input data must be a list() or Spreadsheet().")

        # If Spreadsheet is passed in data, set Spreadsheet to columns
        if spreadsheet_type:
            self.toColumns()
            data.toColumns()

        spreadsheet = list()
        append = spreadsheet.append

        # Loop through the data to create a new spreadsheet
        for row in data:
            append(row)

        # Assign this spreadsheet to the new one
        self.spreadsheet = spreadsheet
        self.initialized = True
        self.transposed = asRows

    def savePath(self, savePath=None):
        """
            Sets or gets the location for saved files

            Attributes:
                savePath (str): location to store saved files
        """
        if savePath is None:
            return self.savePath
        else:
            self.savePath = savePath

    def sort(self, column=0, reverse=False, hasTitle=True):
        """
            Sorts entire spreadsheet based on column

            Attributes:
                column (int): column to sort by
                reverse (bool): reverse sort
                hasTitle (bool): if True, sort spreadsheet from 2nd row (i==1)
        """

        def get_cell(row):
            """
                Returns cell value at column index specified

                Attributes:
                    row (list): row as a list
                    column (int): column index from parent method

                Returns:
                    int, str: value of cell at specified column
            """
            return row[column]

        column = self.getColumnIndex(column)
        self.toRows()

        # Define different parts of the spreadsheet for processing
        spreadsheet = self.spreadsheet
        spreadsheet_header = [spreadsheet[0]]
        spreadsheet_body = spreadsheet[1:]

        # If hasTitle is True, the header is not included in sorting
        if hasTitle is True:
            body = sorted(spreadsheet_body, key=get_cell, reverse=reverse)
            self.spreadsheet = spreadsheet_header + body

        # Include header in sorting
        else:
            self.spreadsheet = sorted(spreadsheet, key=get_cell,
                                      reverse=reverse)

    def toColumns(self):
        """
            Transpose to columns
        """
        # Transpose spreadsheet to edit column
        if not self.getState():
            self.transpose()

    def toRows(self):
        """
            Transpose to rows
        """
        # Transpose spreadsheet to edit rows
        if self.getState():
            self.transpose()

    @preserve_transpose
    def toString(self, fileToString=None):
        """
            Print input to screen

            Attributes:
                fileToString (list): file to print out as string to screen.
        """

        state = self.getState()
        string = ""

        if fileToString is None:
            self.toRows()
            fileToString = self.spreadsheet

        if self.initialized:
            join = str.join

            # Loop through the lines to join them as strings
            for line in fileToString:

                line = [str(item).rjust(20, ' ') for item in line]
                string += join("\t\t", line) + "\n"
        self.revertTranspose(state)
        return string

    def transpose(self):
        """
            Transposes this spreadsheet's rows and columns
        """
        temp_spreadsheet = list()

        # CALL THESE JUST ONCE BEFORE LOOP(S)
        append = temp_spreadsheet.append
        longest_list = len(max(self.spreadsheet, key=len))

        # Loop through the longest row (transposed=f) or column (transposed=t)
        for index in xrange(longest_list):

            # At this index, insert a list for the new row/column
            append(list())

            # CALL THESE JUST ONCE BEFORE LOOP(S)
            append2 = temp_spreadsheet[index].append

            # Loop through current spreadsheet to transpose rows<==>columns
            for line in self.spreadsheet:
                try:
                    append2(line[index])
                except(IndexError):

                    # If the specified cell does not exist, i.e., blank
                    append2("")

        self.spreadsheet = temp_spreadsheet
        self.transposed = not self.getState()

    def revertTranspose(self, prior_state):
        """
        Transpose class if prior_state is different from current state

        Args:
            state1 (bool): original state of document
        """
        self.refresh()
        if prior_state != self.getState():
            self.transpose()
