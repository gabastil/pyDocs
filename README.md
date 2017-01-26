# **pyDocs Package**<a id="top"></a>
---


The pyDocs package is a collection of classes representing different types of documents allowing for programmatic access, manipulation, and output of their data. The classes included in this package are listed below with links to further details for each class.

1. [Spreadsheet](#spreadsheet): allows users to create and manipulate data tables 	 as spreadsheets.
2. [Document](#document): allows users to create, search, clean, and manipulate texts.
3. [HTMLStripper](#htmlstripper): inherits [Document](#document) and strips 	 HTML documents of their HTML tags.

---

## **Spreadsheet** <a id="spreadsheet"></a> ([top](#top))

The Spreadsheet class creates a table data structure (a.k.a. [spreadsheet][def_spreadsheet]), which is a list of lists. The default orientation of this spreadsheet is by row, which means that each element list in the greater list represents a row. This can be changed by transposing the class, turning each element list in the greater list a column representation. The data elements within each row are automatically assigned to the correct row or column as appropriate.

The Spreadsheet class is constructed using the following syntax:

~~~python
import Spreadsheet

# if the source file is tab separated, which is the default delimiter setting
spreadsheet = Spreadsheet(filePath="path.to.tsv")

# if the source file is separated by other characters, e.g., comma ','
spreadsheet = Spreadsheet(filePath="path.to.csv", delimiter=',') 
~~~

The following methods allow users and other classes to access and manipulate the data in the Spreadsheet class:

1. [Methods for Rows](#rows)
2. [Methods for Columns](#columns)
3. [Methods for Data Variables](#data_variables)
4. [Other methods](#other) 

<a id="rows"></a>

###### Methods for Rows

The following methods allow the user and other classes to add, access, change, remove, or get row information for the spreadsheet in Spreadsheet.



<a id="columns"></a>

###### Methods for Columns

<a id="data_variables"></a>

###### Methods for Data Variables

<a id="other"></a>

###### Other Methods


## **Document** <a id="document"></a> ([top](#top))
## **HTMLStripper** <a id="htmlstripper"></a> ([top](#top))

[def_spreadsheet]: README.md "A data set with data contained in rows or columns. Each piece of data can be accessed by two integers, a row number and a column number. For example, the method cell(2,4) retrieves the value in row 2, cell 4."