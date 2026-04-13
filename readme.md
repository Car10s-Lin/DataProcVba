# DataProcVba by catDev
*A collection of modular VBA utilities for cleaning, normalizing, and processing real-world operational data into common forms of operational tables.*

## Framework

### Common

Basic functions that almost all other subs call. Import on default before importing anything else.

- ThisWB[PropertyGet]
- Today[PropertyGet]
- LastRow
- LastCol
- NameToRowIndex
- NameToColIndex
- ArrProc
- NumericEasterEgg

### DataImporting

Import data from online or local snaps. Recommended to call on events like Workbook_Open or Worksheet_Change to allow updating.

- ImportCsv
- GSheetLinkParser
- ListFiles

### DataFiltering
###DataNormalization
###UXControl
