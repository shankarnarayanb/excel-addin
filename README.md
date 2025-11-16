# 🛡️ PII Data Sanitizer (VBA)

**Version 2.4 - Refined Anonymization**

A powerful Excel VBA tool designed for the secure anonymization of Personal Identifiable Information (PII) within Excel and CSV data files. It replaces direct identifiers with realistically formatted, fake data, ensuring compliance with data privacy regulations like GDPR while maintaining data structure for testing and analysis.

> **Important:** This version strictly focuses on direct identifiers and excludes sensitive quasi-identifiers such as GENDER, RACE, and STATUS from sanitization/redaction.

## ✨ Key Features

- **Targeted Anonymization**: Focuses on over 80 variations of direct PII column headers (Names, Emails, Phone Numbers, Addresses, IDs, Financial Info)
- **Realistic Faker Data**: Generates UK-format fake data (names, emails, postcodes, NINOs, etc.) to replace originals, preserving data format and context
- **Secure Deletion**: The original unsanitized file is permanently removed using a 3-pass secure deletion protocol after successful sanitization
- **Preview Mode**: Create a sanitized copy for review before committing to the secure deletion of the original
- **Manual Column Selection**: Override automatic PII detection by selecting columns by Name or Index
- **CSV Import Utility**: Built-in function to quickly import CSV files directly into Excel


## 📖 User Guide

### 1. Prerequisites

- **Microsoft Excel**: The tool is designed to run within the Excel application (VBA environment)
- **Macro-Enabled File**: Your working file (or the file containing the VBA code) must be saved as a Macro-Enabled Workbook (.xlsm) and macros must be enabled
- **VBA Code**: The provided VBA code must be correctly loaded into an Excel module

### 2. PII Detection

The tool operates in two modes for column identification: **Automatic** and **Manual**.

#### 2.1. Automatic Mode (Default)

The tool automatically scans the header row (Row 1) for over 80 column name variations related to PII.

| PII Category | Example Headers Detected |
|--------------|--------------------------|
| **Names** | First Name, Surname, LNAME, Full Name |
| **Contact/Login** | Email, Phone, Mobile, USERNAME, PASSWORD |
| **Address** | Street, Postcode, ZIP, City, Address |
| **National IDs** | NINO, SSN, Passport, Driver's Licence |
| **Financial** | AccountNo, IBAN, SortCode, Credit Card |
| **Digital/Location** | IP Address, LAT, LONG, GPS |

> ℹ️ **Note**: The tool excludes sanitization of GENDER, RACE, and STATUS columns, focusing only on direct personal identifiers.

#### 2.2. Manual Column Selection (Select Columns Button)

Use this when automatic detection is insufficient, or if you only want to sanitize a subset of columns.

1. Click the **Select Columns** button
2. An input box will appear
3. Enter column identifiers separated by a comma. You must choose to use **Names** or **Indices**, but not both:
   - **By Name**: Enter the exact header names (e.g., `Customer ID, Email Address, Last Name`)
   - **By Index**: Enter the column numbers (e.g., `1, 3, 7`)
4. After submission, the tool proceeds directly to the full sanitization process

## 🚀 Usage

### Core Workflow

#### Step 1: Open Data
Open the Excel (.xlsx, .xlsm) or CSV file you wish to sanitize.

#### Step 2: 🔬 Preview (Recommended)
Run the **Preview** function. This creates a new file named `[OriginalName]_preview.[ext]` containing the sanitized data.

**To Preview:**
1. Open the data file
2. Click the **Preview** button
3. The tool creates a copy of your current file in the same directory, appending `_preview` to the filename (e.g., `data.xlsx` becomes `data_preview.xlsx`)
4. The preview file is sanitized and saved. Your original file remains untouched
5. A message box will confirm the location of the preview file

#### Step 3: 💥 Sanitize (IRREVERSIBLE)
Run the **Sanitize File** function. This will:
- Sanitize the current workbook in memory
- Save the sanitized data to a temporary file `[OriginalName]_sanitized.[ext]`
- Securely delete the original file
- Rename the sanitized file back to the original filename
- Open the newly sanitized file

**To Sanitize:**
1. Open the data file (or use the one loaded after a successful Preview)
2. Click the **Sanitize File** button
3. A **CRITICAL WARNING** message box will appear, confirming the irreversible nature of the secure deletion
4. Click **Yes** to proceed
5. The tool performs the sanitization
6. The original file is subjected to a 3-pass secure deletion and then removed from the file system
7. The sanitized copy is renamed to the original filename
8. A success message will display the number of records processed and the time elapsed. The newly sanitized file will be opened

> ⚠️ **WARNING**: The Sanitize File process includes secure 3-pass deletion of the original file. This action is **IRREVERSIBLE**. Always use the Preview function first!

## 🔧 Other Utilities

- **Import CSV** (Import CSV Button): Opens a standard file dialog allowing you to select and directly open a CSV file into a new Excel workbook
- **Help** (Help Button): Displays a quick summary of the workflow and features
- **About** (About Button): Displays the version number and key features

## ⚠️ Important Safety Notes

1. **Always Preview First**: Before running the full sanitization, use the Preview function to verify that the correct columns are being sanitized
2. **Backup Your Data**: Keep backups of original files before running the sanitization process
3. **Test on Sample Data**: Test the tool on a small sample dataset before processing production data
4. **Verify Results**: After sanitization, verify that the data structure and relationships are maintained correctly
5. **Irreversible Process**: The 3-pass secure deletion cannot be undone. Original files are permanently destroyed

## 📋 System Requirements

- Microsoft Excel (Windows)
- VBA/Macro support enabled
- Sufficient permissions to read, write, and delete files in the working directory

## 🔒 Data Privacy Compliance

This tool is designed to help organizations comply with data privacy regulations such as:
- GDPR (General Data Protection Regulation)
- Data Protection Act
- Other privacy frameworks requiring PII anonymization

By replacing direct identifiers with realistic fake data, the tool maintains data utility for testing and analysis while protecting individual privacy.

## 📝 Version History

**Version 2.4 - Refined Anonymization**
- Focus on direct identifiers only
- Exclusion of GENDER, RACE, and STATUS from sanitization
- Enhanced PII detection with 80+ column variations
- 3-pass secure deletion protocol
- Preview mode implementation

---

**License**: 
This project is licensed under the **MIT License** with a **Mandatory Attribution Clause**.

In addition to the standard MIT terms, any use, modification, or redistribution of this software (or its derivatives) must include clear and conspicuous attribution to the original author.

**Support**: For issues, questions, or feature requests, please contact the developer.
