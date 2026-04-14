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
- UiFreeze
- AppCalc
- AppWait
- EnumToIndex
- OpenURL
- NumericEasterEgg

### SnapshotManage

Import, update, or clear data from online or local snapshots. Recommended to call on events like Workbook_Open or Worksheet_Change to allow updating.

- ImportCsv
- GSheetLinkParser
- ListFiles
- OmisDataImporter
- OmisDataPurger
- ClearWs
- PersonalDataProtection
- SoloVal

### DataManipulation

Manipulate data in Range

- AppMatch
- BulkAutoFilter
- FillTitle
- GenSequence
- GetSn
- GetWs
  
### DataNormalization

- BamaToDate
- 
### UIUXControl

- ClearUserForm
- ColWidthCalib
- ColWidthRetrive
- DrawStandardBorders

### TMRTStationInfo

Taichung Metro station info from public data 

- GrSta
- GrStaIndex
