"""Fill dummy data into CRS_Template_Revised_2.xlsx and save as CRS_Template_Dummy.xlsx"""

import sys
import openpyxl

INPUT  = "CRS_Template_Revised_2.xlsx"
OUTPUT = "CRS_Template_Dummy.xlsx"

# Data rows start at Excel row 8 (1-indexed).
# Column A ("Record 1") is already in the template; data goes from column B onward.
DATA_START_ROW = 8

def col_index(sheet, field_name):
    """Return 1-based column index for a field name (row 1 headers)."""
    for cell in sheet[1]:
        if cell.value == field_name:
            return cell.column
    raise KeyError(f"Column '{field_name}' not found in sheet '{sheet.title}'")

def write_row(sheet, excel_row, data: dict):
    """Write a dict of {field_name: value} into a specific Excel row."""
    for field, value in data.items():
        if value is None:
            continue
        col = col_index(sheet, field)
        sheet.cell(row=excel_row, column=col, value=value)

wb = openpyxl.load_workbook(INPUT)

# ── MessageHeader ──────────────────────────────────────────────────────────────
mh = wb["MessageHeader"]
write_row(mh, DATA_START_ROW, {
    "sendingCompanyIN":   "MY12345678",
    "transmittingCountry": "MY",
    "receivingCountry":   "GB",
    "messageType":        "CRS",
    "contact":            "John Doe, +60123456789",
    "messageRefID":       "MY2024GB000000001",
    "messageTypeIndic":   "CRS701",
    "reportingPeriod":    "2024-12-31",
    "timestamp":          "2025-01-15T09:45:30",
})

# ── ReportingFI ────────────────────────────────────────────────────────────────
fi = wb["ReportingFI"]
write_row(fi, DATA_START_ROW, {
    "docTypeIndic":        "OECD1",
    "docRefID":            "MYfi001xyz",
    "resCountryCode":      "MY",
    "in":                  "GIIN123456.00000.LE.458",
    "in_issuedBy":         "MY",
    "in_INType":           "GIIN",
    "name":                "Maybank Berhad",
    "nameType":            "OECD207",
    "address_countryCode": "MY",
    "legalAddressType":    "OECD304",
    "street":              "Jalan Ampang",
    "buildingIdentifier":  "Menara Maybank",
    "postCode":            "50450",
    "city":                "Kuala Lumpur",
    "countrySubentity":    "Wilayah Persekutuan",
})

# ── Individual ─────────────────────────────────────────────────────────────────
ind = wb["Individual"]
individuals = [
    {
        "docTypeIndic":        "OECD1",
        "docRefID":            "MYind001xyz",
        "accountNumber":       "ACC123456789",
        "acctNumberType":      "OECD601",
        "undocumentedAccount": "false",
        "closedAccount":       "false",
        "dormantAccount":      "false",
        "selfCert":            "CRS901",
        "resCountryCode":      "GB",
        "tin":                 "IG12345678010",
        "tin_issuedBy":        "GB",
        "nameType":            "OECD207",
        "title":               "Mr",
        "firstName":           "Ahmad",
        "lastName":            "Abdullah",
        "address_countryCode": "GB",
        "legalAddressType":    "OECD302",
        "street":              "10 Downing Street",
        "districtName":        "Westminster",
        "postCode":            "SW1A 2AA",
        "city":                "London",
        "countrySubentity":    "England",
        "birthDate":           "1980-05-15",
        "birthCity":           "Kuala Lumpur",
        "birthCitySubentity":  "Wilayah Persekutuan",
        "birthCountryCode":    "MY",
        "accountBalance":      "15000.00",
        "currCode":            "MYR",
        "accountType":         "CRS1101",
        "ddProcedure":         "CRS1201",
        "jointAccount":        "false",
    },
    {
        "docTypeIndic":        "OECD1",
        "docRefID":            "MYind002xyz",
        "accountNumber":       "ACC987654321",
        "acctNumberType":      "OECD601",
        "undocumentedAccount": "false",
        "closedAccount":       "false",
        "dormantAccount":      "false",
        "selfCert":            "CRS901",
        "resCountryCode":      "AU",
        "tin":                 "AU98765432",
        "tin_issuedBy":        "AU",
        "nameType":            "OECD207",
        "title":               "Ms",
        "firstName":           "Sarah",
        "lastName":            "Johnson",
        "address_countryCode": "AU",
        "legalAddressType":    "OECD302",
        "street":              "25 Martin Place",
        "postCode":            "2000",
        "city":                "Sydney",
        "countrySubentity":    "New South Wales",
        "birthDate":           "1975-08-20",
        "birthCity":           "Melbourne",
        "birthCitySubentity":  "Victoria",
        "birthCountryCode":    "AU",
        "accountBalance":      "50000.00",
        "currCode":            "USD",
        "accountType":         "CRS1102",
        "ddProcedure":         "CRS1202",
        "jointAccount":        "false",
    },
]
for i, record in enumerate(individuals):
    write_row(ind, DATA_START_ROW + i, record)

# ── Organisation ───────────────────────────────────────────────────────────────
org = wb["Organisation"]
write_row(org, DATA_START_ROW, {
    "docTypeIndic":        "OECD1",
    "docRefID":            "MYorg001xyz",
    "accountNumber":       "CORP987654321",
    "acctNumberType":      "OECD602",
    "undocumentedAccount": "false",
    "closedAccount":       "false",
    "dormantAccount":      "false",
    "selfCert":            "CRS902",
    "acctHolderType":      "CRS101",
    "resCountryCode":      "SG",
    "in":                  "GIIN789012.00000.LE.702",
    "in_issuedBy":         "SG",
    "in_INType":           "GIIN",
    "name":                "ABC Holdings Pte Ltd",
    "nameType":            "OECD207",
    "address_countryCode": "SG",
    "legalAddressType":    "OECD303",
    "street":              "1 Raffles Place",
    "buildingIdentifier":  "One Raffles Place",
    "postCode":            "048616",
    "city":                "Singapore",
    "accountBalance":      "500000.00",
    "currCode":            "SGD",
    "accountType":         "CRS1104",
    "ddProcedure":         "CRS1202",
    "jointAccount":        "false",
})

# ── ControllingPerson ──────────────────────────────────────────────────────────
cp = wb["ControllingPerson"]
write_row(cp, DATA_START_ROW, {
    "accountNumber":       "CORP987654321",
    "ctrlgPersonType":     "CRS801",
    "selfCert":            "CRS1001",
    "resCountryCode":      "SG",
    "tin":                 "SG12345678A",
    "tin_issuedBy":        "SG",
    "nameType":            "OECD207",
    "title":               "Mr",
    "firstName":           "Rajesh",
    "middleName":          "Kumar",
    "lastName":            "Krishnamurthy",
    "address_countryCode": "SG",
    "legalAddressType":    "OECD302",
    "street":              "10 Orchard Road",
    "postCode":            "238843",
    "city":                "Singapore",
    "birthDate":           "1970-03-15",
    "birthCity":           "Singapore",
    "birthCountryCode":    "SG",
})

# ── Payment ────────────────────────────────────────────────────────────────────
pmt = wb["Payment"]
payments = [
    {"accountNumber": "ACC123456789",  "type": "CRS502", "paymentAmnt": "1500.00",  "currCode": "MYR"},
    {"accountNumber": "ACC123456789",  "type": "CRS501", "paymentAmnt": "2500.00",  "currCode": "MYR"},
    {"accountNumber": "ACC987654321",  "type": "CRS501", "paymentAmnt": "5000.00",  "currCode": "USD"},
    {"accountNumber": "CORP987654321", "type": "CRS503", "paymentAmnt": "75000.00", "currCode": "SGD"},
]
for i, record in enumerate(payments):
    write_row(pmt, DATA_START_ROW + i, record)

wb.save(OUTPUT)
print(f"Saved: {OUTPUT}")
