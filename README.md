# Nephthis-BP-VBO---Extended-Excel
Adds functionality missing from the OOB Excel VBO. Due to how Blue Prism has coded instance handle management for this application, this should replace usage of the base VBO, as you cannot mix and match. All actions from the OOB VBO are included in this release. Documentation will only cover new actions which have been added ontop of it.

##  Open Workbook With Password
Opens a workbook using the instance specified in the given handle, opening the book represented by the file name using the password provided.

### Inputs:
* handle - Number
* File name - Text
* Password - Password

### Outputs:
* Workbook Name - Text

## Get Worksheet Range As Collection Offset
Gets the current worksheet into a collection. This will read the worksheet and store the contents into the collection. The offset function allows a starting cell to be specified.

### Inputs:
* handle - Number
* Workbook Name - Text
* Worksheet Name - Text
* StartCell - Text
* Use Header - Flag

### Outputs:
* Data - Collection

## Add Comment to Cell
Adds a comment to a specified cell

### Inputs:
* handle - Number
* cellref - Text
* commentvalue - Text


## Read Comment from Cell
Reads the comment on a specified cell

### Inputs:
* handle - Number
* cellref - Text

### Outputs:
* commentvalue - Text

## Insert Row Below
Adds a row below the current row in Excel

### Inputs:
* Handle - Number
* Workbook - Text
* Worksheet - Text

## Insert Column
Adds a column on either side of current column in Excel

### Inputs:
* Handle - Number
* Workbook - Text
* Worksheet - Text
* On Left - Flag

## Hide Other Workshseets
Hide all other worksheets so that they are not visible.

### Inputs:
* Handle - Number
* Workbook - Text
* Worksheet - Text

## Import CSV
Import CSV worksheets into a workbook using a path

### Inputs:
* Destination Handle - Number
* Destination Workbook - Text
* Destination Worksheet - Text
* Source File Path - Text
* Source Text Qualifier - Text

## Find Relative Cell
Finds a cell and returns a value from a relative location based on fuzzy match

### Inputs:
* Handle - Number
* Text - Text
* Col Offset - Number
* Row Offset - Number

### Outputs:
* Found - Flag
* Cell Value - Text
* Row Number - Number
* Column Number - Number

## Find Relative Cell - Exact Match
Finds a cell and returns a value from a relative location based on exact match

### Inputs:
* Handle - Number
* Text - Text
* Col Offset - Number
* Row Offset - Number

### Outputs:
* Found - Flag
* Cell Value - Text
* Row Number - Number
* Column Number - Number

## Delete Value by Range
Deletes all data from a specified cell range

### Inputs:
* Handle - Number
* Source Worksheet - Text
* Source Workbook - Text
* Source Range - Text

## Clear Content Value by Range
Deletes values only from a specified cell range

### Inputs:
* Handle - Number
* Source Worksheet - Text
* Source Workbook - Text
* Source Range - Text

## Clear Format by Range
Deletes format only from a specified cell range

### Inputs:
* Handle - Number
* Source Worksheet - Text
* Source Workbook - Text
* Source Range - Text

## Change Background Color
Sets background and border colors of specified cell range

### Inputs:
* Handle - Number
* Source Workbook - Text
* Source Worksheet - Text
* Source Range - Text
* Red - Number
* Green - Number
* Blue - Number
* Border Color - Flag
* Border Red - Number
* Border Green - Number
* Border Blue - Number

## Clear Comment Value
Removes comments from specified cell range in current workbook/worksheet

### Inputs:
* Handle - Number
* Source Range - Text

## Select Worksheet Range
Selects a cell range from a specified workbook/worksheet

### Inputs:
* Handle - Number
* Source Workbook - Text
* Source Worksheet - Text
* Cell Range - Text

## Check Cell Color
Retrieves the color of the currently selected cell

### Inputs:
* Handle - Number
* Source Workbook - Text
* Source Worksheet - Text

## Save Workbook As Other Filetype
Allows SaveAs operations for non-xlsx filetypes, as defined by ints from https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlfileformat-enumeration-excel

### Inputs:
* handle - Number
* Workbook Name - Text
* Filename - Text
* Filetype - Text

### Outputs:
* New Workbook Name - Text

## Set Formula
Applies a formula to a specified cell range

### Inputs:
* Handle - Number
* Source Workbook - Text
* Source Worksheet - Text
* Cell Range - Text
* Formula - Text

## Filter Column
Applies up to four excel filters against a given column

### Inputs:
* Handle - Number
* Source Workbook - Text
* Source Worksheet - Text
* Column Header - Text
* Filter - Text
* Filter2 - Text
* Filter3 - Text
* Filter4 - Text
