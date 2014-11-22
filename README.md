# SheetParser #

_SheetParser_ is a simple deserialization system for transforming tabular data into .NET objects, originally used to store lists of items, monsters, etc. in game asset directories. It supports both attribute-driven and procedural transformation of table rows into objects, with the ability to define custom transforms for your own scalar data types.

SheetParser is largely LINQ-based and lazily evaluated. It is provided as a Portable Class Library. It is provided under the MIT license.

## Supported Document Formats ##

- **Excel 2003 XML** (SpreadsheetML)

## Usage ##
Sample code can be found in _EdCanHack.SheetParser.Tests_ for more advanced usage--iterating and simple deserialization can be found there. Test coverage is not 100%, but the project from which this was extracted has been in use for a couple of years.

## Future Work ##
Note: this work is not currently planned, but allows for future possibilities.

- Office OpenXML support
- LibreOffice Calc support
- CSV support
- Relational support - allow an `INamedSheetReader` to deserialize objects with `IEnumerable<T>` members stored in other sheets
