# CRS Excel-to-XML Converter

Converts a filled CRS template Excel file into one or more **FC XML Schema v2.2** files for submission to the **Vanuatu MDES portal** (Ministry of Finance, Vanuatu).

The converter follows the CRS (Common Reporting Standard) data model defined in the OECD Amended CRS XML Schema User Guide v4.0 (October 2024), wrapped in the FC v2.2 envelope format required by the MDES portal.

---

## What This Does

The Vanuatu MDES portal requires CRS reporting in **FC XML Schema v2.2** format. This converter reads your CRS data from a structured Excel template and produces valid FC v2.2 XML files ready for upload.

For large submissions, the output can be split into multiple XML files with balanced row counts.

---

## Data Files

Excel data files are not stored in this repository. They are available on SharePoint:

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

> **Note:** Excel data files (`.xlsx`) and generated output folders (`output_xml*/`) are excluded from version control via `.gitignore`.

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

1. **XML file(s)** — One or more FC XML Schema v2.2 files ready for upload to the Vanuatu MDES portal.
2. **Excel with filepath column** — A copy of the input Excel with a `filepath` column added to `Individual`, `Organisation`, `ControllingPerson`, and `Payment` sheets, showing which XML file each row was written to.

---

## XML Namespace Structure

The generated XML uses three namespaces as required by the MDES portal:

| Prefix | Namespace URI | Used for |
|--------|--------------|---------|
| `ns0` | `urn:fatcacrs:ties:v2` | FC v2.2 wrapper (`FATCA_CRS`, `MessageHeader`, `MessageBody`) |
| `ns1` | `urn:oecd:ties:fatcacrstypes:v2` | All main content elements |
| `ns2` | `urn:oecd:ties:stffatcatypes:v2` | Address children and Name children |

---

## MessageRefId Format

Each XML file is assigned a unique `MessageRefId` generated as:

```
{TransmittingCountry}2025{SendingCompanyIN}{12-char random hex}{part suffix}
```

Example (split file 1 of 50):
```
VU2025562188A3F9C12D4E8BP01
```

| Part | Example | Source |
|------|---------|--------|
| `VU` | TransmittingCountry | MessageHeader sheet |
| `2025` | Reporting year | Fixed |
| `562188` | SendingCompanyIN | MessageHeader sheet |
| `A3F9C12D4E8B` | 12-char random hex | Auto-generated |
| `P01` | Part number | Only added when splitting |

If a `MessageRefId` is provided in the Excel and splitting is not used, the provided value is used as-is.

---

## Splitting Logic

- Splitting unit = account rows (`Individual` + `Organisation` combined).
- Rows are distributed using ceiling division — no file has more than one extra row compared to any other.
- Each split file gets a unique `MessageRefId`.
- The `ReportingFI` block is replicated in every split file.
- `ControllingPerson` and `Payment` rows follow their parent account into the same split file.

---

## Nil Return

Set `MessageTypeIndic` to `CRS703` in the `MessageHeader` sheet to generate a nil-return XML.

---

## Key Notes

- In FC v2.2, both Individual and Organisation accounts use `TIN` for identification numbers (`IN` is not used).
- `AcctNumberType` attribute is not supported in FC v2.2 and is omitted from `AccountNumber`.
- `AddressFix` is preferred over `AddressFree` per MDES portal guidance. Submissions with empty `AddressFix` fields (Street, BuildingIdentifier, PostCode, City) may be flagged for review by the Competent Authority.
- Reference spec: OECD Amended CRS XML Schema User Guide v4.0, October 2024 (`reference/Amended Common Reporting Standard XML Schema.pdf`).
