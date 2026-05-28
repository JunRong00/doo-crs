# CRS Excel-to-XML Converter

Converts a filled CRS template Excel file into **FC XML Schema v2.2** files for submission to the **Vanuatu MDES portal** (Ministry of Finance, Vanuatu).

**Authors:** Kevin Mun (SH) · Jareld Lim (JR)

---

## Table of Contents

1. [Overview](#overview)
2. [Quick Start](#quick-start)
3. [Data Files](#data-files)
4. [Prerequisites](#prerequisites)
5. [Repository Structure](#repository-structure)
6. [Template Structure](#template-structure)
7. [Usage](#usage)
8. [Outputs](#outputs)
9. [Submission Workflow](#submission-workflow)
10. [Technical Reference](#technical-reference)
11. [Key Notes](#key-notes)

---

## Overview

The Vanuatu MDES portal requires CRS (Common Reporting Standard) reporting in **FC XML Schema v2.2** format. This converter reads CRS data from a structured Excel template and produces valid FC v2.2 XML files ready for upload.

- For large submissions, output can be split into multiple balanced XML files
- Each file gets a unique `MessageRefId`
- A copy of the Excel is produced showing which XML file each row was written to

> **Reference:** OECD Amended CRS XML Schema User Guide v4.0, October 2024

---

## Quick Start

```bash
# Single XML file
/opt/anaconda3/bin/python xlsx_to_xml_converter.py YourFile.xlsx

# Split into 50 XML files, saved to output_xml/
/opt/anaconda3/bin/python xlsx_to_xml_converter.py YourFile.xlsx --split 50 --out output_xml
```

---

## Data Files

Excel data files are not stored in this repository due to sensitive financial data. They are available on SharePoint:

> **[CRS Data Files — SharePoint](YOUR_SHAREPOINT_LINK_HERE)**

| File | Description |
|------|-------------|
| `CRS_20260528.xlsx` | CRS submission data |
| `CRS_20260528_v2.xlsx` | CRS submission data (revised) |

---

## Prerequisites

- Python 3.8+
- `pandas`
- `openpyxl`

```bash
pip install pandas openpyxl
```

> **macOS with Anaconda:** Use `/opt/anaconda3/bin/python` — it has the required packages pre-installed.

---

## Repository Structure

```
CRS/
├── xlsx_to_xml_converter.py       # Main converter script
├── README.md                      # This file
├── CLAUDE.md                      # Project notes for Claude Code
├── .gitignore
└── reference/                     # Reference documents (do not modify)
    ├── Amended Common Reporting Standard XML Schema.pdf
    └── xml-schema-crs/            # CRS XML Schema v3.0 XSD files
        ├── CrsXML_v3.0.xsd
        ├── oecdcrstypes_v5.0.xsd
        ├── CommonTypesFatcaCrs_v2.0.xsd
        ├── FatcaTypes_v1.2.xsd
        └── isocrstypes_v1.1.xsd
```

> Excel data files (`.xlsx`) and generated output folders (`output_xml*/`) are excluded from version control.

---

## Template Structure

The Excel template has **6 sheets**:

| Sheet | Purpose |
|-------|---------|
| `MessageHeader` | Message-level metadata (sender, receiver, reporting period, message type) |
| `ReportingFI` | Reporting Financial Institution details |
| `Individual` | Individual account holder records |
| `Organisation` | Entity/organisation account holder records |
| `ControllingPerson` | Controlling persons linked to organisation accounts |
| `Payment` | Payment records linked to individual or organisation accounts |

**Layout convention:**
- Row 1: Column headers
- Rows 2–6: Metadata / remarks (auto-skipped by converter)
- Row 7: Blank separator
- Row 8+: Data records

`ControllingPerson` and `Payment` rows are linked to their parent account via the `accountNumber` column (exact string match).

---

## Usage

### Basic conversion (single XML file)

```bash
/opt/anaconda3/bin/python xlsx_to_xml_converter.py YourFile.xlsx
```

Output: `YourFile_CRS.xml` + `YourFile_with_filepath.xlsx`

### Split into multiple XML files

```bash
/opt/anaconda3/bin/python xlsx_to_xml_converter.py YourFile.xlsx --split 50
```

Output: `YourFile_CRS_part01.xml` … `YourFile_CRS_part50.xml`

### Custom output folder

```bash
/opt/anaconda3/bin/python xlsx_to_xml_converter.py YourFile.xlsx --split 50 --out output_xml
```

### Full options

```
usage: xlsx_to_xml_converter.py [-h] [--split N] [--out FOLDER] input

Arguments:
  input             Path to the filled CRS Excel template (.xlsx)
  --split N, -s N   Split output into N XML files (default: 1)
  --out FOLDER      Output folder for XML files (created if missing)
```

---

## Outputs

| Output | Description |
|--------|-------------|
| `YourFile_CRS.xml` (or `_part01.xml` etc.) | FC v2.2 XML file(s) ready for MDES portal upload |
| `YourFile_with_filepath.xlsx` | Copy of input Excel with `filepath` column showing which XML file each row belongs to |

---

## Submission Workflow

1. **Fill in the Excel template** with CRS data across all 6 sheets
2. **Run the converter** to generate XML files
3. **Review the `_with_filepath.xlsx`** to verify all rows were processed
4. **Log in to the Vanuatu MDES portal** at [https://mdes.doft.gov.vu/MDES/](https://mdes.doft.gov.vu/MDES/)
5. **Upload each XML file** one at a time
6. **Review portal warnings** — incomplete `AddressFix` fields (Street, BuildingIdentifier, PostCode, City) will be flagged for Competent Authority review but do not block submission
7. **Click Next** to confirm the upload

> For large datasets, split into 50 files and upload each part separately.

---

## Technical Reference

### XML Namespace Structure

| Prefix | Namespace URI | Used for |
|--------|--------------|---------|
| `ns0` | `urn:fatcacrs:ties:v2` | FC v2.2 wrapper (`FATCA_CRS`, `MessageHeader`, `MessageBody`) |
| `ns1` | `urn:oecd:ties:fatcacrstypes:v2` | All main content elements |
| `ns2` | `urn:oecd:ties:stffatcatypes:v2` | Address children and Name children |

### MessageRefId Format

Each XML file is assigned a unique `MessageRefId`:

```
{TransmittingCountry}2025{SendingCompanyIN}{12-char random hex}{part suffix}
```

Example — split file 1 of 50: `VU2025562188A3F9C12D4E8BP01`

| Part | Example | Source |
|------|---------|--------|
| `VU` | TransmittingCountry | MessageHeader sheet |
| `2025` | Reporting year | Fixed |
| `562188` | SendingCompanyIN | MessageHeader sheet |
| `A3F9C12D4E8B` | 12-char random hex | Auto-generated |
| `P01` | Part number | Only added when splitting |

### Splitting Logic

- Splitting unit = account rows (`Individual` + `Organisation` combined)
- Rows distributed using ceiling division — no file has more than one extra row
- Each split file gets a unique `MessageRefId`
- The `ReportingFI` block is replicated in every split file
- `ControllingPerson` and `Payment` rows follow their parent account

### Nil Return

Set `MessageTypeIndic` to `CRS703` in the `MessageHeader` sheet to generate a nil-return XML.

---

## Key Notes

- In FC v2.2, both Individual and Organisation accounts use `TIN` for identification numbers (`IN` is not used)
- `AcctNumberType` attribute is omitted from `AccountNumber` (not supported in FC v2.2)
- `AddressFix` is preferred over `AddressFree` per MDES portal guidance
- The converter is built against the OECD CRS XML Schema v3.0 data model, wrapped in the FC v2.2 envelope required by the Vanuatu MDES portal
