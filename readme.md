# DataProcVBA by catDev

*A modular VBA toolkit for cleaning, normalizing, and transforming messy real-world operational data into structured, usable tables.*

## Overview

This project is a collection of reusable VBA modules designed for real-world operational environments, where data is often inconsistent, unstructured, and manually maintained.

Instead of treating Excel as a simple spreadsheet tool, this framework approaches it as a lightweight data processing system under constrained conditions.

## Architecture

The toolkit is organized into modular layers, each handling a specific stage of the data workflow:

Data Source

   ↓
   
SnapshotManage   (data ingestion & refresh)

   ↓
   
DataFilter       (filtering & selection)

   ↓
   
TableConstructor (table structuring)

   ↓
   
DataLookup       (reference & matching)

   ↓
   
DataNormalization (format standardization)

   ↓
   
UIControl        (presentation & interaction)

## Example Use Case

## Framework

### Common

Core utility functions shared across all modules.

Must be loaded before other components.

- `ThisWB`[PropertyGet]
   - Behaviour
      - Returns `ThisWorkbook`
   - Purpose
      - Provides a shorter alias for `ThisWorkbook` to simplify code readability and reduce repetitive typing.
- `Today`[PropertyGet]
   - Behaviour
      - Returns the current system date as a `Date`.
   - Purpose
      - Provides a simplified interface for retrieving the current date, especially for users transitioning from Excel formulas.
- `LastRow(ws As Worksheet, Optional indexCol As Long = 1, Optional chaoMode As Boolean = False, Optional searchForBlank As Boolean = False) As Long`
   - Behavior  
      - Returns the last row index of a worksheet.
      - Mininum result is `1`
   - Purpose  
      - Detects the boundary of data within a worksheet, useful for determining dataset size, locating the next available row for writing, or identifying layout limits in Excel-based documents.
   - Parameters  
      - `ws`: Worksheet. Required. The target worksheet.  
      - `indexCol`: Long. Optional. Column used to determine the last row when `chaoMode = False`. Defaults to column 1.  
      - `chaoMode`: Boolean. Optional. If set to `True`, performs a full worksheet scan and returns the largest row index containing any data. Useful for messy or irregular tables.  
      - `searchForBlank`: Boolean. Optional. If set to `True`, adds `1` to the result, allowing detection of the first available row for writing.
- `LastCol(ws As Worksheet, Optional indexRow As Long = 1, Optional chaoMode As Boolean = False, Optional searchForBlank As Boolean = False) As Long`
   - Behavior  
      - Returns the last column index of a worksheet.
      - Mininum result is `1`
   - Purpose  
      - Detects the horizontal boundary of data within a worksheet, useful for determining dataset width, locating the next available column for writing, or identifying layout limits in Excel-based documents.
   - Parameters  
      - `ws`: Worksheet. Required. The target worksheet.  
      - `indexRow`: Long. Optional. Row used to determine the last column when `chaoMode = False`. Defaults to row 1.  
      - `chaoMode`: Boolean. Optional. If set to `True`, performs a full worksheet scan and returns the largest column index containing any data. Useful for messy or irregular tables.  
      - `searchForBlank`: Boolean. Optional. If set to `True`, adds `1` to the result, allowing detection of the first available column for writing.
- `NameToRowIndex(nm As String, ws As Worksheet, Optional smart As Boolean = True, Optional startRow As Long = 1, Optional indexCol As Long = 1) As Long`
   - Behavior  
      - Returns the row index of `nm` within a specified column.
      - Returns `-1` if `nm` is not found.
   - Purpose  
      - Resolves row positions based on row names or key values, enabling dynamic row referencing in functions such as `Worksheet.Cells()`, `LastCol()`, and `BulkAutoFilter()`.
   - Parameters  
      - `nm`: String. Required. The target value or row name to search for.  
      - `ws`: Worksheet. Required. The target worksheet.  
      - `smart`: Boolean. Optional. Defaults to `True`. When enabled, performs a fuzzy match using `Like` with wildcard patterns.  
      - `startRow`: Long. Optional. Starting row for the search. Defaults to row 1.  
      - `indexCol`: Long. Optional. Column used for lookup. Defaults to column 1.
- `NameToColIndex(nm As String, ws As Worksheet, Optional smart As Boolean = True, Optional startCol As Long = 1, Optional indexRow As Long = 1) As Long`
   - Behavior  
      - Returns the column index of `nm` within a specified row.
      - Returns `-1` if `nm` is not found.
   - Purpose  
      - Resolves column positions based on header names or key values, enabling dynamic column referencing in functions such as `Worksheet.Cells()`, `LastRow()`, and `BulkAutoFilter()`.
   - Parameters  
      - `nm`: String. Required. The target value or column name to search for.  
      - `ws`: Worksheet. Required. The target worksheet.  
      - `smart`: Boolean. Optional. Defaults to `True`. When enabled, performs a fuzzy match using `Like` with wildcard patterns.  
      - `startCol`: Long. Optional. Starting column for the search. Defaults to column 1.  
      - `indexRow`: Long. Optional. Row used for lookup. Defaults to row 1.
- `ArrProc(arr As Variant, Optional ipt As Variant, Optional method As ArrProcMethod = arrAppend) As Variant`
   - Behavior  
      - Performs array operations based on the specified `method`.
      - Returns the processed array as a variant as a new array.
      - Supports structural modifications, transformations, parsing, and sorting of arrays through a unified interface.
   - Purpose  
      - Provides a single entry point for common array manipulation tasks in VBA, reducing the need for multiple specialized functions.
      - Designed for flexible and composable operations in real-world data processing workflows.
   - Parameters  
      - `arr`: Variant. Required. The target array.  
      - `ipt`: Variant. Optional. Overloaded parameter. Its meaning depends on the selected `method` (value, index, array, or delimiter). See methods below for expected ipt type and content.
      - `method`: ArrProcMethod. Optional. Determines the operation to perform. Defaults to `arrAppend`.
   - Methods  
      - **arrAppend**  
         - Appends `ipt` to the end of the array.
         - `ipt`: Variant. Required.
      - **arrPop**  
         - Removes the last element of the array.
         - `ipt`: Not used.
      - **arrInsert**  
         - Inserts an element (`ipt(1)`) before a specified position (`ipt(0)`).  
         - `ipt`: Array. Required. → `(index, value)`
      - **arrRemove**  
         - Removes the element at the specified index.
         - `ipt`: Long. Required.
      - **arrSortAscending / arrSortDescending**  
         - Sorts the array.  
         - Supports mixed types by separating numeric and string values.
         - `ipt`: Not used.
      - **arrReverse**  
         - Reverses the array.
         - `ipt`: Not used.
      - **arrConcat**  
         - Flattens and concatenates nested arrays into a single array.
         - `ipt`: Not used.
      - **arrDivide**  
         - Splits the array into two parts at a given index (`ipt`).
         - `ipt`: Long. Required.
      - **arrSplit**  
         - Splits a string into an array using a delimiter.
         - Uses `", "` as the default delimiter.
         - Wraps VBA's built-in `Split()` for simpler usage.
         - `ipt`: String. Optional.
      - **arrSplitChar**  
         - Splits a string into an array of individual characters.
         - `arr`: Array. Required. Dim an empty array before calling this method.
         - `ipt`: String. Required.
      - **arrJoin**  
         - Wraps VBA's built-in `Join()` function.
         - `ipt`: String. Required.
   - Example  

     ```
      arr = Array(3, 1, 2)

      arr = ArrProc(arr, , arrSortAscending)
       → Array(1, 2, 3)
      
      arr = ArrProc(arr, 4, arrAppend)
       → Array(1, 2, 3, 4)
      
      arr = ArrProc(arr, Array(2, 99), arrInsert)
      ' → Array(1, 2, 99, 3, 4)
     ```
      
- `UiFreeze(opt As UiFreezeOption)`
   - Behavior  
      - Toggles Excel application-level UI and calculation settings, including:
         - `EnableEvents`
         - `DisplayAlerts`
         - `ScreenUpdating`
         - `Calculation`
         - `StatusBar`
      - Controlled through `UiFreezeOption` (`uiActivate` / `uiDeactivate`).
   - Purpose  
      - Temporarily suspends UI updates and automatic calculations to improve performance and prevent unintended side effects during macro execution.
      - Particularly useful in workbooks that combine VBA procedures with formula-based calculations.
   - Notes  
      - Should be used with care, as disabling application-level settings may affect workbook behavior if not properly restored.
   - Parameters  
      - `opt`: UiFreezeOption. Required.  
         - `uiActivate`: Enables UI and automatic calculation  
         - `uiDeactivate`: Disables UI updates and switches calculation mode
- `AppCalc()`
   - Behavior  
      - Forces workbook-wide calculation and waits until all calculations are completed.
   - Purpose  
      - Ensures data consistency in workbooks that combine VBA procedures with formula-based calculations.
- `AppWait(Optional waitSec As Long = 0)`
   - Behavior  
      - Pauses execution using `Application.Wait` for the specified number of seconds.
   - Purpose  
      - Provides a simple wrapper for introducing delays in VBA execution.
   - Parameters  
      - `waitSec`: Long. Optional.  
         - Number of seconds to wait  
         - If not specified or 0, no delay is applied
- `EnumToIndex(en As Long, Optional ofst As Long = 0) As Long`
   - Behavior  
      - Converts a `2^n`-style enum value into its corresponding index `n`.
   - Purpose
      - Enables enum values to be used as array indices or for more readable mapping logic.
   - Example

     ```
     Enum fruits
        fruitApple = 2 ^ 0
        fruitOrange = 2 ^ 1
        fruitBanana = 2 ^ 2
     End Enum
     
     Function ReturnFruit(en as fruits) As String
        Dim a As Array
        a = Array("apple", "orange", "banana")
        idx = EnumToIndex(en)
        ReturnFruit = a(idx)
     End Function`
     
     ' ReturnFruit(fruitOrange) → "orange"
     ```
     
   - Parameters
      - `en`: Long. Required. Enum value to convert.
      - `ofst`: Long. Optional. Offset added to the resulting index.
- `OpenURL(url As String)`
   - Behavior  
      - Opens default browser with `url`.
   - Purpose
      - Provides a flexible way to trigger external links from VBA, including UserForms and ActiveX controls.
   - Parameters
      - `url`: String. Required. Url to be opened.
- `GetWs(Optional idx As wsIndex = wsNone, Optional wsName As String = "") As Worksheet`
   - Behavior  
      - Returns a `Worksheet` object based on either a predefined enum (`wsIndex`) or a worksheet name.
   - Purpose  
      - Simplifies worksheet access and reduces repetitive referencing logic in VBA code.
   - Parameters  
      - `idx`: wsIndex. Optional. Enum representing a worksheet.  
         - `wsIndex` must be predefined as a `2^n`-style enum before use.  
      - `wsName`: String. Optional. Name of the worksheet.
- `NumericEasterEggs(numInput As Long, Optional allowVandalism As Boolean = False)`
   - Behavior  
      - Triggers hidden numeric easter eggs based on the value of `numInput`.
      - Depending on the input, may display messages, open external URLs, or perform workbook-level modifications.
      - Uses an internal trigger guard to prevent repeated execution within the same call chain.
   - Purpose  
      - Provides a collection of developer-side hidden behaviors for humor, experimentation, and limited internal interaction.
      - Intended as a non-essential utility and should not be used in production-facing workflows.
   - Parameters  
      - `numInput`: Long. Required. Numeric trigger used to activate a specific easter egg.  
      - `allowVandalism`: Boolean. Optional. Defaults to `False`. When set to `True`, enables destructive or workbook-altering behaviors for selected triggers.
   - Notes  
      - Some triggers are harmless and only display messages or open links.
      - Some triggers may modify workbook content, close the workbook without saving, or alter worksheet formatting when `allowVandalism = True`.
      - Not recommended for use in serious or production workbooks.

### SnapshotManage

Handles data ingestion from local and online sources (e.g., CSV, Google Sheets), including refresh and cleanup.

- ImportCsv
- GSheetLinkParser
- ListFiles
- OmisDataImporter
- OmisDataPurger
- ClearWs
- PersonalDataProtection
- SoloVal

### DataFilter

Provides filtering and statistical selection tools for extracting relevant subsets from datasets.

- BulkAutoFilter
- NthQuantile

### TableConstructor

Constructs structured tables using column definitions and primary keys.

Designed for practical use cases where Excel serves as a lightweight data system (~30k rows).

- FillTitle
- GenSequence
- GetSn

### DataLookup

Enhanced lookup utilities that extend beyond built-in VBA functions, offering more flexible and robust matching logic.

- `AppMatch(search As Variant, within As Variant, Optional matchOption As AppMatchOption = matchExactly, Optional xmatchMode As Boolean = False) As Long`
   - Behavior  
      - Returns the location (1-based) of first encounter of `search` in `within` like regular `Application.Match`
      - `matchOption` allows find biggest, smallest, and exactly.
      - When `xmatchMode = True`, `within` will be turned backwards for starting search from last element.
      - Returns `-1` if encounters error.
   - Purpose  
      - Provides a simple wrapper for `Application.Match`.
      - Simulates `XMATCH` on older version of Excel
      - Provides fallback on error for match function. Returns `-1` as sentinel for error finding to allow easy processing.
   - Parameters  
      - `search`: Variant. Required. Items to look for in `within`.  
      - `within`: Variant. Required. 1-D array or range to search within.  
      - `matchOption`: AppMatchOption. Optional. Match method. Uses `matchExactly` if not assigned.
         - matchGreater: Search for the smallest value that's greater than `search` 
         - matchMinor: Search for the biggest value that's minor than `search` 
         - matchExactly: Search for the exact item same with `search`
      - `xmatchMode`: Boolean. Optional. When set to `True`, looks for `search` backwards in `within`, then returns the location of first encounter. 
   - Example  
      - `SmartMid("ABCDEFG", 3, 2)` → `"CD"`  
      - `SmartMid("ABCDEFG", 5, -2)` → `"DE"`  
      - `SmartMid("ABCDEFG", 4, 0)` → `"DEFG"`
- `SmartMid(inputStr As String, Optional start As Long = 1, Optional length As Long = 0)`
   - Behavior  
      - Returns a substring from `inputStr`, extending VBA's built-in `Mid()` behavior.
      - When `length > 0`, behaves like the standard `Mid()` function.
      - When `length < 0`, extracts characters in reverse direction from the given start position.
      - When `length = 0`, returns the substring from `start` to the end of the string.
      - If reverse extraction would cause an invalid position, falls back to `Mid(inputStr, start)`.
   - Purpose  
      - Provides a more flexible substring function for cases where standard `Mid()` is too limited, especially when backward extraction or simplified end-of-string slicing is needed.
   - Parameters  
      - `inputStr`: String. Required. The source string to extract from.  
      - `start`: Long. Optional. Starting position for extraction. Defaults to `1`.  
      - `length`: Long. Optional. Length of substring to extract.  
         - Positive: extract forward  
         - Negative: extract backward  
         - Zero: extract from `start` to the end of the string
   - Example  
      - `SmartMid("ABCDEFG", 3, 2)` → `"CD"`  
      - `SmartMid("ABCDEFG", 5, -2)` → `"DE"`  
      - `SmartMid("ABCDEFG", 4, 0)` → `"DEFG"`

### DataNormalization

Standardize inconsistent data formats (e.g., dates, numeric representations) into structured and uniform forms.

- BamaToDate
- SmartNumFormat

### UIControl

Utilities for managing worksheet and userform interactions, including UI cleanup, layout adjustments, and visual consistency.

- ClearUserForm
- ColWidthCalib
- ColWidthRetrive
- DrawStandardBorders

### TMRTStationInfo

Domain-specific station mapping utilities for Taichung Metro operational forms and tables.

This public version includes only non-sensitive, publicly accessible reference data adapted for demonstration purposes.

- TmrtGreenLineStations[PropertyGet]
- GrSta
- GrStaIndex

## Notes
