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

## Framework

### Common

Core utility functions shared across all modules.

Must be loaded before other components.

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
- GetWs
- NumericEasterEgg

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

- AppMatch
- SmartMid

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

- GrSta
- GrStaIndex
- TmrtGreenLineStations[PropertyGet]

## Usage

## Example Use Case

## Notes
