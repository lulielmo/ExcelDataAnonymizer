# Excel Data Anonymizer for Finance and Operations User Reports

A specialized Python tool for anonymizing personal data in Finance and Operations user license reports while preserving the original formatting. This tool is specifically designed and optimized for handling "Rapport över antal Finance and Operations-användare" Excel files.

## Optimized For
This tool is specifically designed to handle:
- Finance and Operations user license reports
- Reports containing user information such as aliases, usernames, and email addresses
- Excel files with the specific structure of F&O user reports
- Standard system roles and license types that should remain unchanged

## Features
- Anonymizes names, email addresses, and usernames in F&O user reports
- Maintains consistent mapping between original and anonymized values
- Preserves Excel formatting (fonts, colors, borders, etc.)
- Handles standard F&O values that should not be anonymized (e.g., Access License Type, System user, Security Role)
- Supports both simple anonymization and format-preserving anonymization

## Installation

1. Install Python 3.8 or later
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Anonymization
To anonymize a Finance and Operations user report:

```bash
python anonymize_excel.py SysUserLicenseCountReport.xlsx anonymized_report.xlsx
```

### Format-Preserving Anonymization
To anonymize while preserving the original Excel formatting:

```bash
python preserve_formatting.py SysUserLicenseCountReport.xlsx anonymized_report.xlsx anonymized_report.mapping.json formatted_report.xlsx
```

### Arguments
- `input.xlsx`: Path to the F&O user report Excel file
- `anonymized_output.xlsx`: Path where the anonymized file will be saved
- `anonymized.mapping.json`: Mapping file containing original-to-anonymized value pairs
- `formatted_output.xlsx`: Path where the formatted and anonymized file will be saved

## Output Files
The scripts generate the following files:
1. Anonymized Excel file (`anonymized_output.xlsx`)
2. JSON mapping file (`anonymized_output.mapping.json`) containing:
   - Original names to anonymized names
   - Original email addresses to anonymized ones
   - Original usernames to anonymized ones
3. Formatted anonymized file (`formatted_output.xlsx`) - when using preserve_formatting.py

## Protected Values
The tool automatically preserves F&O system values that should not be anonymized, such as:
- Access License Type
- System user
- Security Role
- Teammedlemmar
- Mobility user
- Other standard F&O system values

## Example Usage

1. Basic anonymization of a user report:
```bash
python anonymize_excel.py SysUserLicenseCountReport.xlsx anonymized_report.xlsx
```

2. Format-preserving anonymization:
```bash
python preserve_formatting.py SysUserLicenseCountReport.xlsx anonymized_report.xlsx anonymized_report.mapping.json formatted_report.xlsx
```

## Notes
- The tool is specifically designed for Finance and Operations user license reports
- Maintains consistency in anonymization across all occurrences of the same value
- Email domains are preserved while the local part is anonymized
- All formatting from the source Excel file is preserved in the formatted output
- The mapping file allows for tracking the relationship between original and anonymized values

## Important Note
This tool is specifically optimized for Finance and Operations user license reports ("Rapport över antal Finance and Operations-användare"). While it may work with other Excel files, its functionality is primarily designed and tested for this specific report type.
