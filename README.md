# HMC Scanner Reader

This folder contains tools to extract and analyze HMC (Hardware Management Console) scanner data from Power Systems and generate comprehensive reports.

## Overview

The HMC Scanner Reader extracts information from Excel files (`.xls` and `.xlsx`) in the `HMCscannerfile` directory and creates detailed Word document reports with tables showing:

- System Information (hostname, serial number, model, firmware)
- LPAR (Logical Partition) Information
- Processor Information
- Memory Information
- Network Adapters (excluding fiber cards as requested)

## Files

- `extract_hmc_report.py` - Main script to extract data and generate Word report
- `requirements.txt` - Python dependencies
- `HMCscannerfile/` - Directory containing Excel scanner data files
- `README.md` - This documentation file

## Prerequisites

1. Python 3.6 or higher
2. Required Python packages (install using pip):
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the Report Generator

```bash
python extract_hmc_report.py
```

The script will:
- Look for Excel files (`.xls` and `.xlsx`) in the `HMCscannerfile/` directory
- Extract system information from each Excel file and its sheets
- Generate a Word document report named `System_Server_Report.docx`

### 3. Output

The script generates a Word document with the title "System Server Report" containing the following sections:

1. **Summary** - Overview of total systems, LPARs, processors, memory, and network adapters
2. **System Information** - Table with server details
3. **LPAR Information** - Table with LPAR details
4. **Processor Information** - Table with processor details
5. **Memory Information** - Table with memory details
6. **Network Adapters** - Table with network adapter details (excluding fiber cards)

## Data Sources

The script processes Excel files from HMC scanner data, which may contain:
- System hardware information
- LPAR configurations
- Network adapter details
- Processor and memory information
- Firmware levels

## Excel File Processing

The script automatically:
- Reads all sheets from Excel files
- Identifies sheet content based on sheet names and column headers
- Extracts relevant data based on content patterns
- Handles both `.xls` and `.xlsx` file formats
- Uses pandas and openpyxl for robust Excel file reading

## Sample Output

The generated report includes tables like:

| Hostname | Serial Number | Model | Firmware | Execution Date |
|----------|---------------|-------|----------|----------------|
| Server1 | ABC123 | Power System | v1.0 | 2025-01-01 |

## Notes

- Fiber cards and ethernet information are excluded from the network adapters table as requested
- The script automatically detects and processes all Excel files in the `HMCscannerfile` directory
- Default status values are used for components where specific status information is not available
- The script intelligently identifies data types based on sheet names and content

## Troubleshooting

If you encounter issues:

1. **No files found**: Ensure Excel files exist in `HMCscannerfile/` directory
2. **Import errors**: Install required packages with `pip install -r requirements.txt`
3. **Permission errors**: Ensure you have write permissions in the current directory
4. **Excel reading errors**: The script uses both pandas and openpyxl as fallbacks

## File Structure

```
hmcscanreader/
├── extract_hmc_report.py    # Main script
├── requirements.txt         # Python dependencies
├── README.md               # This file
├── HMCscannerfile/         # Excel scanner data files
│   └── *.xls              # Excel files with scanner data
└── System_Server_Report.docx # Generated report (after running script)
```

## Dependencies

- `python-docx` - For creating Word documents
- `pandas` - For Excel file reading and data manipulation
- `openpyxl` - For Excel file reading (fallback)
- `xlrd` - For legacy Excel file support
