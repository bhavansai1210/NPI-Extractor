# NPI Register API Data Extraction

## Overview

This Python script is designed to fetch National Provider Identifier (NPI) data using the NPI Registry API provided by the Centers for Medicare & Medicaid Services (CMS). The script takes a list of individuals' first names and last names from an Excel file, queries the NPI Registry API, and updates the Excel file with the corresponding NPI numbers.

## Prerequisites

- Python 3.x
- pandas library
- requests library

Install the required libraries using the following command:

```bash
pip install pandas requests
```

## Usage
Input Data: Place your input Excel file containing first names and last names in the same directory as the script. Ensure the columns are named "firstName" and "lastName".

Run the Script:

```bash
python script_name.py
```

Replace script_name.py with the actual name of your Python script.

## Output:

An output Excel file will be generated with an additional column "NPI" containing the fetched NPI numbers.
The final Excel file is saved to 'output_excel_file.xlsx' in the same directory as the script.


## Resuming from the Last State
The script checks for the existence of an intermediate Excel file ('intermediate_excel_file.xlsx') and a resume information file ('resume_info.txt'). If these files exist, the script will resume from the last saved state.

## Cleaning Up
The intermediate Excel file and the resume information file are removed once the script completes successfully.

## Error Handling
If an error occurs during the API request, the script logs the error and continues processing the next record.
Note
Ensure internet connectivity is available for successful API requests.
## Dependencies
- pandas
- requests
