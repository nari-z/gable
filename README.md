# gable
Tools for incremental management of excel-file. this is possible to manage excel-file in XML format.


# Setup
`go get github.com/nari-z/gable`


# Usage
convert excel-file to xml-file.

`gable export [SourceExcelFilePath] [DestXMLFilePath]`

convert xml-file to excel-file.

`gable import [SourceXMLFilePath] [DestExcelFilePath]`


# Supported Properties
this property is supported.
- cell value
  - date format values are not supported(to be corrected).
- border
- font
- background color  
(Coming soon...)
