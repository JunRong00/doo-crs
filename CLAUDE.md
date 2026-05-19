# CLAUDE.md — CRS Converter Project

## Python Environment

Always use `/opt/anaconda3/bin/python` on this machine — it has `pandas` and `openpyxl` installed. The default `python` command uses the `vqa` conda env which lacks `pandas`.

```bash
/opt/anaconda3/bin/python xlsx_to_xml_converter.py CRS_Template_Revised_2.xlsx
```

---

## Project Overview

This project converts a CRS (Common Reporting Standard) Excel template into CRS XML Schema v3.0 files for submission to tax authorities.

**Reference spec:** OECD Amended CRS XML Schema User Guide v4.0, October 2024 (`doi:10.1787/dd7ee57a-en`)  
**Schema version:** CRS v3.0

---

## Key Files

| File | Purpose |
|------|---------|
| `xlsx_to_xml_converter.py` | Main converter (719 lines) |
| `fill_dummy_data.py` | Test data generator |
| `CRS_Template_Revised_2.xlsx` | Blank data entry template |
| `CRS_Template_Dummy.xlsx` | Pre-filled sample (output of fill_dummy_data.py) |
| `CRS_Template_Dummy_CRS.xml` | Sample XML output for reference |

---

## Template Layout

- Row 1: Column headers
- Rows 2–6: Metadata/remark rows (`REMARK_ROWS = 6`) — skipped by converter
- Row 7: Blank separator
- Row 8+: Data records (`DATA_START_ROW = 8`)

`read_sheet()` skips 6 remark rows and drops column A (the label column) before returning a DataFrame.

---

## XML Namespaces (Appendix B, p.58 of spec)

```python
NS = {
    "crs": "urn:oecd:ties:crs:v3",
    "stf": "urn:oecd:ties:crsstf:v5",
    "cfc": "urn:oecd:ties:commontypesfatcacrs:v2",
    "iso": "urn:oecd:ties:isocrstypes:v1",
}
```

---

## Critical Schema Rules

### SelfCert Enum Types (two different enums — easy to confuse)

| Context | Enum Type | Values |
|---------|-----------|--------|
| AccountHolder (Individual/Organisation) | `CrsSelfCert_EnumType` | CRS901=true, CRS902=false, CRS900=not reported |
| ControllingPerson | `CrsSelfCertforCtrlgPerson_EnumType` | CRS1001=true, CRS1002=false, CRS1000=not reported |

Using CRS901 for ControllingPerson is a schema violation.

### CorrMessageRefId — NOT used in CRS

The spec explicitly marks `CorrMessageRefId` as "Optional (non-CRS)" at both MessageSpec (p.9) and DocSpec (p.26) levels. Do not add it to CRS XML output.  
`corrDocRefId` IS the correct correction reference for CRS.

### Multi-value fields

The schema allows repeating elements that the template handles as single columns:
- `ResCountryCode` (1..∞), `TIN` (0..∞), `Name` (1..∞), `Address` (1..∞), org `IN` (0..∞)

This is acceptable for standard CRS filings. Edge cases requiring multiple values per field are not supported by the current template design.

---

## Converter Architecture

```
convert()
├── read_sheet() × 6 sheets
├── For nil return (CRS703): build_message_spec() only
├── split_rows() — distributes Individual+Organisation rows into N chunks
└── Per chunk:
    ├── build_message_spec()     — unique MessageRefId per split file
    ├── build_reporting_fi()     — replicated in every split file
    └── Per account row:
        ├── build_individual_account()   or
        ├── build_organisation_account()
        │   ├── build_doc_spec()
        │   ├── build_account_number()
        │   ├── build_account_balance()
        │   ├── build_controlling_person() (linked by accountNumber)
        │   └── build_payment() (linked by accountNumber)
        └── write_filepath_column()  — adds filepath column to Excel copy
```

### Why MessageHeader and ReportingFI have no filepath column

Their data appears in **every** split XML file — there is no single filepath to assign. Only `Individual`, `Organisation`, `ControllingPerson`, and `Payment` sheets get the filepath column.

---

## Linking Logic

`ControllingPerson` and `Payment` rows are linked to account rows via `accountNumber` (exact string match). The converter uses `get_linked(df, acc_num)` for this join. Rows with no matching parent account are silently skipped.

---

## Splitting Logic

- Splitting unit = account rows (Individual + Organisation combined).
- `split_rows()` uses `math.ceil(total / n)` — no file has more than one extra row.
- Each split file gets a unique `MessageRefId` generated as `{tc}{year}{rc}P{part:02d}{uuid6}`.
- If `--split N` exceeds the number of account rows, N is capped to the row count.

---

## AccountHolder Element Order (schema-required)

Individual: `EquityInterestType → SelfCert → Individual`  
Organisation: `EquityInterestType → SelfCert → Organisation → AcctHolderType`  
ControllingPerson: `Individual → CtrlgPersonType → SelfCert`

---

## Address Element Order (schema-required, AddressFix)

`Street → BuildingIdentifier → SuiteIdentifier → FloorIdentifier → DistrictName → POB → PostCode → City → CountrySubentity`

If `addressFree` is provided and no structured fields exist, `AddressFree` is used instead of `AddressFix`.
