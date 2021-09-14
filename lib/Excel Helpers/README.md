# VBA Helpers for Excel
A collection of generic modules that make working with Excel objects a little more convivial.

## File Selector
_FileSelector.bas_

Contains the method `ChooseFile`, which shows the user a file selector dialogue configured for Excel files, and returns the selected file's path as a String. Does not actually open the file.

## Sheet Object Helper
_SheetObjectHelper.bas_

Contains a number of methods for creating, finding, and manipulating Worksheets.

## List Object Helprt
_ListObjectHelper.bas_

Contains a number of methods for creating, finding, and manipulating ListObjects (Excel Tables).
### Has Dependencies:
- SheetObjectHelper.bas
- ValidateHeadersResult.cls

## Validate Headers Result
_ValidateHeadersResult.cls_

This defines an object returned by the `ListObjectHelper.ValidateHeaders(TheList As ListObject, HasHeaders As Collection)` method. Instead of a simple Boolean (or an error in case of invalid headers), the method returns this object as a report on whether it found the required headers, including which were found, and which were not found.

- `.HasRequiredHeaders` : __Boolean__ : Indicates whether or not all members of the HasHeaders Collection were found as headers of the ListObject.
- `.ImplementedHeaders ` : __Collection__ : The Collection of all members of the HasHeaders Collection that were found as header column labels in TheList.
- `.MissingHeaders` : __Collection__ : The Collection of all members of the HasHeaders Colllection that were not found as header column labels in TheList.
- `.ExtraHeaders` : __Collection__ : The Collection of the headers in TheList that did not appear in the HasHeaders Collection.

