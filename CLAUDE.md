# CLAUDE.md — CRS Converter Project

## Python Environment

Always use `/opt/anaconda3/bin/python` on this machine — it has `pandas` and `openpyxl` installed.

```bash
/opt/anaconda3/bin/python xlsx_to_xml_converter.py YourFile.xlsx --split 50 --out output_xml
```

---

## Project Overview

Converts a CRS (Common Reporting Standard) Excel template into **FC XML Schema v2.2** files for submission to the **Vanuatu MDES portal**.

**Reference files (local only — do NOT fetch external specs):**
1. `reference/Amended Common Reporting Standard XML Schema.pdf` — OECD CRS XML Schema User Guide v4.0 (Oct 2024)
2. `reference/xml-schema-crs/` — CRS v3.0 XSD files:
   - `CrsXML_v3.0.xsd`
   - `oecdcrstypes_v5.0.xsd`
   - `CommonTypesFatcaCrs_v2.0.xsd`
   - `FatcaTypes_v1.2.xsd`
   - `isocrstypes_v1.1.xsd`

---

## Key Files

| File | Purpose |
|------|---------|
| `xlsx_to_xml_converter.py` | Main converter |
| `reference/` | Reference PDF and XSD files (read-only) |

---

## Template Layout

- Row 1: Column headers
- Rows 2–6: Metadata/remark rows (`REMARK_ROWS = 6`) — skipped by converter
- Row 7: Blank separator
- Row 8+: Data records (`DATA_START_ROW = 8`)

`read_sheet()` skips 6 remark rows and drops column A (the label column) before returning a DataFrame.

---

## XML Namespaces (FC v2.2 — portal requirement)

```python
NS_MAIN  = "urn:fatcacrs:ties:v2"            # ns0 — wrapper only
NS_TYPES = "urn:oecd:ties:fatcacrstypes:v2"  # ns1 — main content elements
NS_STF   = "urn:oecd:ties:stffatcatypes:v2"  # ns2 — Address children + Name children
```

These namespace URIs were discovered through portal validation errors. They are NOT in the local reference files (which describe CRS v3.0).

---

## FC v2.2 vs CRS v3.0 Differences

| Feature | CRS v3.0 (reference files) | FC v2.2 (portal requirement) |
|---------|---------------------------|-------------------------------|
| Root element | `CRS_OECD version="3.0"` | `FATCA_CRS version="2.2"` |
| Message wrapper | `MessageSpec` / `CrsBody` | `MessageHeader` / `MessageBody` |
| MessageType | `"CRS"` | `"FATCA-CRS"` |
| Org identifier | `IN` element | `TIN` element |
| AcctNumberType | Attribute on AccountNumber | Not allowed in FC v2.2 |
| Address children ns | `commontypesfatcacrs:v2` | `stffatcatypes:v2` |
| Name children ns | main CRS namespace | `stffatcatypes:v2` |

---

## MessageRefId Format

```
{TransmittingCountry}2025{SendingCompanyIN}{12-char random hex}{Pxx if split}
```

Example: `VU2025562188A3F9C12D4E8BP01`

---

## Critical Schema Rules

### AccountHolder Element Order

Individual: `Individual` only (no SelfCert, EquityInterestType in FC v2.2)
Organisation: `Organisation` → `AcctHolderTypeCRS`
ControllingPerson: `Individual` → `CtrlgPersonType`

### CorrMessageRefId — NOT used in CRS

The spec marks `CorrMessageRefId` as "Optional (non-CRS)". Do not add it to CRS output.
`CorrDocRefId` IS the correct correction reference for CRS.

### Address Element Order (AddressFix)

`Street → BuildingIdentifier → SuiteIdentifier → FloorIdentifier → DistrictName → POB → PostCode → City → CountrySubentity`

---

## Converter Architecture

```
convert()
├── read_sheet() × 6 sheets
├── For nil return (CRS703): build_message_header() only
├── split_rows() — distributes Individual+Organisation rows into N chunks
└── Per chunk:
    ├── build_message_header()    — unique MessageRefId per split file
    ├── build_reporting_fi()      — replicated in every split file
    └── Per account row:
        ├── build_individual_account()   or
        ├── build_organisation_account()
        │   ├── build_doc_spec()
        │   ├── build_account_number()
        │   ├── build_account_balance()
        │   ├── build_controlling_person()
        │   └── build_payment()
        └── write_filepath_column()
```

## Linking Logic

`ControllingPerson` and `Payment` rows linked to accounts via `accountNumber` (exact string match).

## Splitting Logic

- Unit = account rows (Individual + Organisation combined)
- `math.ceil(total / n)` distribution
- Each split file gets a unique `MessageRefId`
- `--split N` is capped to the row count if N exceeds total rows
