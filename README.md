# Lab Results Parser

A Python tool for parsing laboratory test results from text files and converting them into structured Excel spreadsheets. Designed specifically for BioFire respiratory panel reports and similar lab result formats.

## Overview

This tool extracts key information from lab report text files including:
- Run dates
- Sample/Specimen IDs
- Patient age/sex
- Completion date/time
- Test results (detected/not detected)

The parsed data is exported to Excel format for easy analysis and record-keeping.

## Features

- Parses multiple specimen records from a single text file
- Extracts all detected pathogens from test results
- Handles various specimen ID formats (SPEC #, Specimen:)
- GUI application for easy file selection and processing
- Progress tracking during file processing
- Automatic deduplication of specimen records

## Components

### [parse_lab_results.py](parse_lab_results.py)
Command-line script for parsing lab results. Uses regex pattern matching to extract specimen data and test results from text files.

### [lab_report_extractor.py](lab_report_extractor.py)
Tkinter-based GUI application that provides a simple interface for file selection and extraction.

### [LabParserApp/lab_results_app.py](LabParserApp/lab_results_app.py)
PyQt5-based GUI application with enhanced features:
- Threaded processing for better responsiveness
- Progress bar for long-running operations
- Input/output file selection dialogs
- Error handling and user feedback

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd lab-results-parser
```

2. Install required dependencies:
```bash
pip install pandas openpyxl PyQt5
```

## Usage

### Command Line

Run the script directly on a text file:
```bash
python parse_lab_results.py
```
Edit the `file_path` variable in the script to point to your input file.

### GUI Application (Tkinter)

```bash
python lab_report_extractor.py
```

### GUI Application (PyQt5)

```bash
python LabParserApp/lab_results_app.py
```

Then follow the on-screen instructions:
1. Select your lab results text file
2. Choose where to save the output Excel file
3. Click "Process File" to extract the data

## Input Format

The tool expects text files containing lab reports with the following fields:
- `RUN DATE:` - The date the test was run
- `SPEC #:` or `Specimen:` - The specimen identifier
- `AGE/SEX:` - Patient demographic information
- `COMP:` - Completion date/time
- Test results with "Final" status and "Detected"/"Not Detected" results

## Output Format

The tool generates an Excel file with the following columns:
- Date
- Sample ID #
- Age
- COMP DATE-Time
- Result (lists all detected pathogens or "Not detected")

## Requirements

- Python 3.6+
- pandas
- openpyxl
- PyQt5 (for GUI application)
- tkinter (usually included with Python)
