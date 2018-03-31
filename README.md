# WorksheetWithHeader
Convert content in a worksheet to data structure.
The worksheet contains a header row to help identify the keywords for the data
There is a Parameter section in the worksheet, which followed the rules below to build:
 - Parameter is a keyword in the header row
 - Columns start from "Parameter" keyword are the section of Parameter
 - It has two kind of format to store the key value pair in Parameter section:
   - Key in the header row and value in same column below the header
   - Key and value are in contiguous cells at same row
