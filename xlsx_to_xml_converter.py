# python xlsx_to_xml_converter.py CRS_Template_Revised_2.xlsx --split 5 --out C:/submissions/CRS_2024

#!/usr/bin/env python3
"""
CRS Excel to XML Converter  v3.0
Converts a CRS template Excel file to one or more CRS XML Schema v3.0 files.

Reference: Amended Common Reporting Standard XML Schema User Guide v4.0
           OECD, October 2024 (doi:10.1787/dd7ee57a-en)

Usage:
    python crs_excel_to_xml.py <input.xlsx> [--split N] [--out <prefix>]

Arguments:
    input.xlsx       Path to the filled CRS Excel template.
    --split N        Split into N XML files with balanced rows. Default: 1.
    --out <prefix>   Output filename prefix. Default: same stem as input.
                     N=1  -> <prefix>_CRS.xml
                     N>1  -> <prefix>_CRS_part01.xml, <prefix>_CRS_part02.xml ...

Outputs:
    1. One or more CRS XML files.
    2. A copy of the Excel with a new 'filepath' column on each account sheet
       (Individual, Organisation, ControllingPerson, Payment) showing which
       XML file each row was assigned to.

Splitting logic:
    - Splitting unit = account rows (Individual + Organisation combined).
    - Rows are distributed using ceiling division so no file has more than
      one extra row compared to any other file.
    - Each split file has its own unique MessageRefId.
    - ReportingFI block is identical in every split file (always required).
    - ControllingPerson and Payment rows travel with their parent account.
"""

import sys
import math
import uuid
import shutil
import argparse
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.dom import minidom

# ── Namespaces (Appendix B, p.58) ─────────────────────────────────────────────
NS = {
    "crs": "urn:oecd:ties:crs:v3",
    "stf": "urn:oecd:ties:crsstf:v5",
    "cfc": "urn:oecd:ties:commontypesfatcacrs:v2",
    "iso": "urn:oecd:ties:isocrstypes:v1",
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

def crs(tag): return f"{{{NS['crs']}}}{tag}"
def stf(tag): return f"{{{NS['stf']}}}{tag}"
def cfc(tag): return f"{{{NS['cfc']}}}{tag}"

# ── Template layout ────────────────────────────────────────────────────────────
# After pd.read_excel(header=0), DataFrame rows are:
#   0=Element/Attribute  1=Requirement  2=Size
#   3=Input Type         4=Example      5=Separator (blank)
#   6+ = actual data records
REMARK_ROWS    = 6
DATA_START_ROW = 8   # Excel row number where data begins (1-based)

# ── Helpers ────────────────────────────────────────────────────────────────────

def safe(val):
    """Return stripped string or None for blank/NaN."""
    if val is None:
        return None
    if isinstance(val, float) and pd.isna(val):
        return None
    s = str(val).strip()
    return s if s else None

def sub(parent, tag, text=None, attrib=None):
    el = ET.SubElement(parent, tag, attrib or {})
    if text is not None:
        el.text = text
    return el

def uid(cc):
    return f"{cc}{uuid.uuid4().hex[:12].upper()}"

def bool_attr(val):
    if val is None:
        return None
    return "true" if str(val).strip().lower() in ("true", "1", "yes") else "false"

def fmt_date(val):
    if not val:
        return None
    try:
        return pd.to_datetime(val).strftime("%Y-%m-%d")
    except Exception:
        return str(val).strip()

def fmt_amount(val):
    if val is None:
        return "0.00"
    try:
        return f"{float(val):.2f}"
    except Exception:
        return "0.00"

def read_sheet(xl_dict, name):
    """
    Read a sheet by name (case-insensitive).
    Skips the 6 remark rows, drops label column A,
    lowercases column names, removes fully-empty rows.
    """
    key = next((k for k in xl_dict if k.strip().lower() == name.lower()), None)
    if key is None:
        return pd.DataFrame()
    df = xl_dict[key]
    if len(df) <= REMARK_ROWS:
        return pd.DataFrame()
    df = df.iloc[REMARK_ROWS:].reset_index(drop=True)
    if df.shape[1] > 0:
        df = df.iloc[:, 1:]          # drop col A label
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.where(pd.notna(df), None)
    df = df.dropna(how="all").reset_index(drop=True)
    return df

def row_dict(row):
    return {str(k).strip().lower(): v for k, v in row.items()}

def split_rows(df, n):
    """
    Distribute df rows into n balanced chunks using ceiling division.
    Max difference between any two chunks is at most 1 row.
    """
    total = len(df)
    if total == 0 or n <= 1:
        return [df]
    chunk_size = math.ceil(total / n)
    return [df.iloc[i:i + chunk_size].reset_index(drop=True)
            for i in range(0, total, chunk_size)]

def write_xml(root, path):
    raw    = ET.tostring(root, encoding="unicode")
    pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding="UTF-8")
    with open(path, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        for line in pretty.decode("utf-8").splitlines():
            if not line.startswith("<?xml"):
                f.write(line + "\n")

# ── XML element builders ───────────────────────────────────────────────────────

def build_doc_spec(parent, doc_type, doc_ref, corr_message_ref=None, corr_doc_ref=None):
    """DocSpec_Type. Child names use CRS v3.0 schema casing."""
    ds = sub(parent, stf("DocSpec"))
    sub(ds, stf("DocTypeIndic"), doc_type or "OECD1")
    sub(ds, stf("DocRefId"),     doc_ref)
    if corr_message_ref:
        sub(ds, stf("CorrMessageRefId"), corr_message_ref)
    if corr_doc_ref:
        sub(ds, stf("CorrDocRefId"), corr_doc_ref)
    return ds

def address_country(row):
    return (safe(row.get("address_countrycode")) or
            safe(row.get("countrycode")) or
            safe(row.get("rescountrycode")))

def build_res_country(parent, row, required=False, context="party"):
    country = (safe(row.get("rescountrycode")) or
               safe(row.get("address_countrycode")) or
               safe(row.get("countrycode")))
    if country:
        sub(parent, crs("ResCountryCode"), country)
        return country
    if required:
        raise ValueError(f"{context} is missing required ResCountryCode.")
    return None

def build_address(parent, row, required=False, context="party"):
    """
    Address_Type — Section IId, p.13-14.
    legalAddressType = Attribute on Address element.
    AddressFix order: Street, BuildingIdentifier, SuiteIdentifier,
    FloorIdentifier, DistrictName, POB, PostCode, City (Validation), CountrySubentity.
    """
    country   = address_country(row)
    legal_type= safe(row.get("legaladdresstype"))
    addr_free = safe(row.get("addressfree"))
    street    = safe(row.get("street"))
    bldg      = safe(row.get("buildingidentifier"))
    suite     = safe(row.get("suiteidentifier"))
    floor_id  = safe(row.get("flooridentifier"))
    district  = safe(row.get("districtname"))
    pob       = safe(row.get("pob"))
    post_code = safe(row.get("postcode"))
    city      = safe(row.get("city"))
    subentity = safe(row.get("countrysubentity"))

    if not country and required:
        raise ValueError(f"{context} is missing required address CountryCode.")
    if not country:
        return None

    attrib = {"legalAddressType": legal_type} if legal_type else {}
    addr_el = ET.SubElement(parent, cfc("Address"), attrib)
    sub(addr_el, cfc("CountryCode"), country)

    if addr_free and not any([street, bldg, city, post_code]):
        sub(addr_el, cfc("AddressFree"), addr_free)
    else:
        fix = sub(addr_el, cfc("AddressFix"))
        if street:    sub(fix, cfc("Street"),             street)
        if bldg:      sub(fix, cfc("BuildingIdentifier"), bldg)
        if suite:     sub(fix, cfc("SuiteIdentifier"),    suite)
        if floor_id:  sub(fix, cfc("FloorIdentifier"),    floor_id)
        if district:  sub(fix, cfc("DistrictName"),       district)
        if pob:       sub(fix, cfc("POB"),                pob)
        if post_code: sub(fix, cfc("PostCode"),           post_code)
        sub(fix, cfc("City"), city or "UNKNOWN")
        if subentity: sub(fix, cfc("CountrySubentity"),   subentity)
        if addr_free:
            sub(addr_el, cfc("AddressFree"), addr_free)
    return addr_el

def build_person_name(parent, row):
    """
    NamePerson_Type — Section IIc, p.11-13.
    nameType = Attribute. FirstName & LastName = Validation elements.
    Child order: PrecedingTitle, Title, FirstName, MiddleName, NamePrefix,
    LastName, GenerationIdentifier, Suffix, GeneralSuffix.
    """
    name_type  = safe(row.get("nametype"))
    attrib = {"nameType": name_type} if name_type else {}
    name_el = ET.SubElement(parent, crs("Name"), attrib)

    v = lambda k: safe(row.get(k))
    if v("precedingtitle"):       sub(name_el, crs("PrecedingTitle"),       v("precedingtitle"))
    if v("title"):                sub(name_el, crs("Title"),                v("title"))
    sub(name_el, crs("FirstName"), v("firstname") or "NFN")
    if v("middlename"):           sub(name_el, crs("MiddleName"),           v("middlename"))
    if v("nameprefix"):           sub(name_el, crs("NamePrefix"),           v("nameprefix"))
    sub(name_el, crs("LastName"), v("lastname") or "UNKNOWN")
    if v("generationidentifier"): sub(name_el, crs("GenerationIdentifier"), v("generationidentifier"))
    if v("suffix"):               sub(name_el, crs("Suffix"),               v("suffix"))
    if v("generalsuffix"):        sub(name_el, crs("GeneralSuffix"),        v("generalsuffix"))
    return name_el

def build_org_name(parent, row):
    """Organisation Name — Section IIIc, p.16. nameType = Attribute."""
    name_type = safe(row.get("nametype"))
    attrib = {"nameType": name_type} if name_type else {}
    el = ET.SubElement(parent, crs("Name"), attrib)
    el.text = safe(row.get("name")) or "UNKNOWN"
    return el

def build_tin(parent, row):
    """TIN — Section IIb, p.11. issuedBy = Attribute on TIN element."""
    tin_val   = safe(row.get("tin"))
    issued_by = safe(row.get("tin_issuedby"))
    if not tin_val:
        return None
    attrib = {"issuedBy": issued_by} if issued_by else {}
    el = ET.SubElement(parent, crs("TIN"), attrib)
    el.text = tin_val
    return el

def build_org_in(parent, row):
    """Organisation IN — Section IIIb, p.16. issuedBy & INType = Attributes."""
    in_val    = safe(row.get("in"))
    issued_by = safe(row.get("in_issuedby"))
    in_type   = safe(row.get("in_intype"))
    if not in_val:
        return None
    attrib = {}
    if issued_by: attrib["issuedBy"] = issued_by
    if in_type:   attrib["INType"]   = in_type
    el = ET.SubElement(parent, crs("IN"), attrib)
    el.text = in_val
    return el

def build_birth_info(parent, row):
    """
    BirthInfo — Section IIf, p.14-15.
    Child elements: BirthDate, City, CitySubentity, CountryInfo.
    """
    birth_date = fmt_date(safe(row.get("birthdate")))
    birth_city = safe(row.get("birthcity"))
    birth_sub  = safe(row.get("birthcitysubentity"))
    birth_cc   = safe(row.get("birthcountrycode"))
    former_cc  = safe(row.get("formercountryname"))

    if not any([birth_date, birth_city, birth_cc, former_cc]):
        return None
    el = sub(parent, crs("BirthInfo"))
    if birth_date: sub(el, crs("BirthDate"),     birth_date)
    if birth_city: sub(el, crs("City"),           birth_city)
    if birth_sub:  sub(el, crs("CitySubentity"),  birth_sub)
    if birth_cc or former_cc:
        ci = sub(el, crs("CountryInfo"))
        if birth_cc:    sub(ci, crs("CountryCode"),      birth_cc)
        elif former_cc: sub(ci, crs("FormerCountryName"), former_cc)
    return el

def build_account_number(parent, row):
    """
    AccountNumber — Section IVd, p.18.
    AcctNumberType, UndocumentedAccount, ClosedAccount, DormantAccount
    are all ATTRIBUTES on AccountNumber element (not child elements).
    """
    acc_num   = safe(row.get("accountnumber")) or "NANUM"
    acct_type = safe(row.get("acctnumbertype"))
    undoc     = bool_attr(safe(row.get("undocumentedaccount")))
    closed    = bool_attr(safe(row.get("closedaccount")))
    dormant   = bool_attr(safe(row.get("dormantaccount")))

    attrib = {}
    if acct_type:         attrib["AcctNumberType"]      = acct_type
    if undoc   == "true": attrib["UndocumentedAccount"] = "true"
    if closed  == "true": attrib["ClosedAccount"]       = "true"
    if dormant == "true": attrib["DormantAccount"]      = "true"

    el = ET.SubElement(parent, crs("AccountNumber"), attrib)
    el.text = acc_num
    return el

def build_account_balance(parent, row):
    """AccountBalance — Section IVg, p.22. currCode = Attribute."""
    curr = safe(row.get("currcode")) or "USD"
    el   = ET.SubElement(parent, crs("AccountBalance"), {"currCode": curr})
    el.text = fmt_amount(safe(row.get("accountbalance")))
    return el

def build_payment(parent, pmt_row):
    """Payment — Section IVh, p.22-23. currCode = Attribute on PaymentAmnt."""
    pmt_type = safe(pmt_row.get("type"))
    if not pmt_type:
        return None
    curr    = safe(pmt_row.get("currcode")) or "USD"
    pmt_el  = sub(parent, crs("Payment"))
    sub(pmt_el, crs("Type"), pmt_type)
    amnt_el = ET.SubElement(pmt_el, crs("PaymentAmnt"), {"currCode": curr})
    amnt_el.text = fmt_amount(safe(pmt_row.get("paymentamnt")))
    return pmt_el

def build_controlling_person(parent, cp_row):
    """
    ControllingPerson — Section IVf, p.21-22.
    Order (schema diagram p.49): Individual -> CtrlgPersonType -> SelfCert.
    """
    cp_el  = sub(parent, crs("ControllingPerson"))
    ind_el = sub(cp_el,  crs("Individual"))

    acc = safe(cp_row.get("accountnumber")) or "unknown account"
    build_res_country(ind_el, cp_row, required=True,
                      context=f"ControllingPerson for {acc}")

    build_tin(ind_el, cp_row)
    build_person_name(ind_el, cp_row)
    build_address(ind_el, cp_row, required=True,
                  context=f"ControllingPerson for {acc}")
    build_birth_info(ind_el, cp_row)

    sub(cp_el, crs("CtrlgPersonType"), safe(cp_row.get("ctrlgpersontype")) or "CRS800")
    sub(cp_el, crs("SelfCert"),        safe(cp_row.get("selfcert"))        or "CRS1000")
    return cp_el

def get_linked(df, acc_num):
    """Return rows from df where accountnumber matches acc_num."""
    if df.empty or not acc_num:
        return pd.DataFrame()
    return df[df["accountnumber"].astype(str).str.strip() == acc_num]

# ── Section builders ───────────────────────────────────────────────────────────

def build_message_spec(root, header_row, tc, part_num=1, total_parts=1):
    """
    MessageSpec — Section I, p.8-10.
    Each split file gets a unique MessageRefId.
    If splitting, Warning field notes the part.
    """
    msg = sub(root, crs("MessageSpec"))

    sending_in = safe(header_row.get("sendingcompanyin"))
    if sending_in:
        sub(msg, crs("SendingCompanyIN"), sending_in)

    rc = safe(header_row.get("receivingcountry")) or "XX"
    sub(msg, crs("TransmittingCountry"), tc)
    sub(msg, crs("ReceivingCountry"),    rc)
    sub(msg, crs("MessageType"),         "CRS")

    warning = safe(header_row.get("warning")) or ""
    if total_parts > 1:
        warning = f"{warning} Split file {part_num} of {total_parts}.".strip()
    if warning:
        sub(msg, crs("Warning"), warning)

    contact = safe(header_row.get("contact"))
    if contact:
        sub(msg, crs("Contact"), contact)

    rep_period = safe(header_row.get("reportingperiod")) or datetime.now().strftime("%Y-%m-%d")
    try:
        year = pd.to_datetime(rep_period).strftime("%Y")
    except Exception:
        year = datetime.now().strftime("%Y")

    # Always generate unique MessageRefId when splitting
    provided = safe(header_row.get("messagerefid"))
    if provided and total_parts == 1:
        msg_ref_id = provided
    else:
        suffix     = f"P{part_num:02d}" if total_parts > 1 else ""
        msg_ref_id = f"{tc}{year}{rc}{suffix}{uuid.uuid4().hex[:6].upper()}"

    sub(msg, crs("MessageRefId"),    msg_ref_id)
    sub(msg, crs("MessageTypeIndic"), safe(header_row.get("messagetypeindic")) or "CRS701")
    corr_msg = safe(header_row.get("corrmessagerefid"))
    if corr_msg:
        sub(msg, crs("CorrMessageRefId"), corr_msg)
    sub(msg, crs("ReportingPeriod"), fmt_date(rep_period) or rep_period)
    sub(msg, crs("Timestamp"),       datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S"))
    return msg

def build_reporting_fi(crs_body, fi_row, tc):
    """
    ReportingFI — Section IVa, p.17.
    Order: ResCountryCode -> IN -> Name -> Address -> DocSpec.
    CRS v3.0 defines ReportingFI as CorrectableOrganisationParty_Type, which
    extends OrganisationParty_Type by appending DocSpec.
    """
    fi_el = sub(crs_body, crs("ReportingFI"))
    build_res_country(fi_el, fi_row, required=False)
    build_org_in(fi_el, fi_row)
    build_org_name(fi_el, fi_row)
    build_address(fi_el, fi_row, required=True, context="ReportingFI")
    build_doc_spec(fi_el,
                   safe(fi_row.get("doctypeindic"))  or "OECD1",
                   safe(fi_row.get("docrefid"))      or uid(tc),
                   safe(fi_row.get("corrmessagerefid")),
                   safe(fi_row.get("corrdocrefid")))
    return fi_el

def build_individual_account(rg_el, rd, cp_df, pmt_df, tc):
    """
    Individual Account Report — Section IVc, p.17.
    AccountHolder order (IVe, p.19): EquityInterestType -> SelfCert -> Individual.
    """
    ar_el = sub(rg_el, crs("AccountReport"))
    build_doc_spec(ar_el,
                   safe(rd.get("doctypeindic"))  or "OECD1",
                   safe(rd.get("docrefid"))      or uid(tc),
                   safe(rd.get("corrmessagerefid")),
                   safe(rd.get("corrdocrefid")))
    build_account_number(ar_el, rd)

    ah_el = sub(ar_el, crs("AccountHolder"))
    equity = safe(rd.get("equityinteresttype"))
    if equity: sub(ah_el, crs("EquityInterestType"), equity)
    sub(ah_el, crs("SelfCert"), safe(rd.get("selfcert")) or "CRS900")

    ind_el = sub(ah_el, crs("Individual"))
    build_res_country(ind_el, rd, required=True,
                      context=f"Individual account {safe(rd.get('accountnumber')) or 'unknown'}")
    build_tin(ind_el, rd)
    build_person_name(ind_el, rd)
    build_address(ind_el, rd, required=True,
                  context=f"Individual account {safe(rd.get('accountnumber')) or 'unknown'}")
    build_birth_info(ind_el, rd)

    acc_num = safe(rd.get("accountnumber"))
    for _, cp_row in get_linked(cp_df, acc_num).iterrows():
        build_controlling_person(ar_el, row_dict(cp_row))

    build_account_balance(ar_el, rd)

    for _, pmt_row in get_linked(pmt_df, acc_num).iterrows():
        build_payment(ar_el, row_dict(pmt_row))

    sub(ar_el, crs("DDProcedure"), safe(rd.get("ddprocedure"))  or "CRS1200")
    sub(ar_el, crs("AccountType"), safe(rd.get("accounttype"))  or "CRS1100")

    joint = safe(rd.get("jointaccount"))
    if joint and joint.lower() == "true":
        ja = sub(ar_el, crs("JointAccount"))
        sub(ja, crs("Number"), safe(rd.get("jointaccount_number")) or "1")

def build_organisation_account(rg_el, rd, cp_df, pmt_df, tc):
    """
    Organisation Account Report — Section IVc, p.17.
    AccountHolder order (IVe, p.19-20):
        EquityInterestType -> SelfCert -> Organisation -> AcctHolderType.
    """
    ar_el = sub(rg_el, crs("AccountReport"))
    build_doc_spec(ar_el,
                   safe(rd.get("doctypeindic"))  or "OECD1",
                   safe(rd.get("docrefid"))      or uid(tc),
                   safe(rd.get("corrmessagerefid")),
                   safe(rd.get("corrdocrefid")))
    build_account_number(ar_el, rd)

    ah_el = sub(ar_el, crs("AccountHolder"))
    equity = safe(rd.get("equityinteresttype"))
    if equity: sub(ah_el, crs("EquityInterestType"), equity)
    sub(ah_el, crs("SelfCert"), safe(rd.get("selfcert")) or "CRS900")

    org_el = sub(ah_el, crs("Organisation"))
    build_res_country(org_el, rd, required=False)
    build_org_in(org_el, rd)
    build_org_name(org_el, rd)
    build_address(org_el, rd, required=True,
                  context=f"Organisation account {safe(rd.get('accountnumber')) or 'unknown'}")

    sub(ah_el, crs("AcctHolderType"), safe(rd.get("acctholdertype")) or "CRS102")

    acc_num = safe(rd.get("accountnumber"))
    for _, cp_row in get_linked(cp_df, acc_num).iterrows():
        build_controlling_person(ar_el, row_dict(cp_row))

    build_account_balance(ar_el, rd)

    for _, pmt_row in get_linked(pmt_df, acc_num).iterrows():
        build_payment(ar_el, row_dict(pmt_row))

    sub(ar_el, crs("DDProcedure"), safe(rd.get("ddprocedure"))  or "CRS1200")
    sub(ar_el, crs("AccountType"), safe(rd.get("accounttype"))  or "CRS1100")

    joint = safe(rd.get("jointaccount"))
    if joint and joint.lower() == "true":
        ja = sub(ar_el, crs("JointAccount"))
        sub(ja, crs("Number"), safe(rd.get("jointaccount_number")) or "1")

# ── Excel filepath column writer ───────────────────────────────────────────────

def write_filepath_column(input_path, output_excel_path,
                          ind_map, org_map, cp_map, pmt_map):
    """
    Copy input Excel to output_excel_path and add a 'filepath' column
    to each account sheet showing which XML file each data row belongs to.

    Maps are {0-based data row index: xml_filepath_string}.
    Data rows start at Excel row DATA_START_ROW (row 8).
    """
    shutil.copy2(input_path, output_excel_path)
    wb = openpyxl.load_workbook(output_excel_path)

    fp_fill  = PatternFill("solid", fgColor="FFF2CC")
    fp_font  = Font(name="Arial", size=10, italic=True, color="7F6000")
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    hdr_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")

    sheet_maps = {
        "Individual":        ind_map,
        "Organisation":      org_map,
        "ControllingPerson": cp_map,
        "Payment":           pmt_map,
    }

    for sheet_name, fp_map in sheet_maps.items():
        if sheet_name not in wb.sheetnames or not fp_map:
            continue
        ws      = wb[sheet_name]
        new_col = ws.max_column + 1
        col_ltr = get_column_letter(new_col)

        # Header
        hdr = ws.cell(row=1, column=new_col, value="filepath")
        hdr.fill = hdr_fill
        hdr.font = hdr_font

        # Blank remark rows 2-7
        for rr in range(2, DATA_START_ROW):
            ws.cell(row=rr, column=new_col, value="")

        # Filepath per data row
        for data_idx, filepath in fp_map.items():
            excel_row = DATA_START_ROW + data_idx
            c = ws.cell(row=excel_row, column=new_col, value=filepath)
            c.fill = fp_fill
            c.font = fp_font

        ws.column_dimensions[col_ltr].width = 45

    wb.save(output_excel_path)

# ── Main converter ─────────────────────────────────────────────────────────────

def convert(input_path, n_splits=1, output_prefix=None):
    print(f"\n{'='*60}")
    print(f"  CRS Excel to XML Converter v3.0")
    print(f"{'='*60}")
    print(f"  Input  : {input_path}")
    print(f"  Splits : {n_splits}")

    xl = pd.read_excel(input_path, sheet_name=None, dtype=str, header=0)

    header_df = read_sheet(xl, "MessageHeader")
    fi_df     = read_sheet(xl, "ReportingFI")
    ind_df    = read_sheet(xl, "Individual")
    org_df    = read_sheet(xl, "Organisation")
    cp_df     = read_sheet(xl, "ControllingPerson")
    pmt_df    = read_sheet(xl, "Payment")

    if header_df.empty:
        raise ValueError("MessageHeader sheet is missing or has no data rows.")

    header_row = row_dict(header_df.iloc[0])
    tc  = safe(header_row.get("transmittingcountry")) or "MY"
    mti = safe(header_row.get("messagetypeindic"))    or "CRS701"

    # Output paths
    base    = Path(input_path).stem if output_prefix is None else output_prefix
    out_dir = Path(input_path).parent

    def xml_path(part, total):
        if total == 1:
            return str(out_dir / f"{base}_CRS.xml")
        return str(out_dir / f"{base}_CRS_part{part:02d}.xml")

    excel_out = str(out_dir / f"{base}_with_filepath.xlsx")

    # ── Nil return ─────────────────────────────────────────────────────────────
    if mti == "CRS703":
        print("  Mode   : Nil return (CRS703) — CrsBody omitted")
        root = ET.Element(crs("CRS_OECD"), {"version": "3.0"})
        build_message_spec(root, header_row, tc)
        op = xml_path(1, 1)
        write_xml(root, op)
        print(f"\n  ✅  {op}")
        return [op], None

    if fi_df.empty:
        raise ValueError("ReportingFI sheet is missing or has no data rows.")
    fi_row = row_dict(fi_df.iloc[0])

    # ── Tag account rows with source sheet and original index ──────────────────
    ind_tagged = ind_df.copy()
    ind_tagged["__sheet__"]    = "Individual"
    ind_tagged["__orig_idx__"] = range(len(ind_tagged))

    org_tagged = org_df.copy()
    org_tagged["__sheet__"]    = "Organisation"
    org_tagged["__orig_idx__"] = range(len(org_tagged))

    all_accounts   = pd.concat([ind_tagged, org_tagged], ignore_index=True)
    total_accounts = len(all_accounts)

    if total_accounts == 0:
        raise ValueError("No account rows found in Individual or Organisation sheets.")

    # Cap splits to available rows
    actual_splits = min(n_splits, total_accounts)
    if actual_splits < n_splits:
        print(f"  ⚠️   Only {total_accounts} account(s) — reducing to {actual_splits} file(s).")

    chunks      = split_rows(all_accounts, actual_splits)
    total_parts = len(chunks)

    print(f"  Accounts: {total_accounts} rows → {total_parts} file(s)")
    print(f"{'='*60}\n")

    # Maps: {0-based data row index in original df -> xml filepath}
    ind_map = {}
    org_map = {}
    cp_map  = {}
    pmt_map = {}
    xml_files = []

    # ── One XML per chunk ──────────────────────────────────────────────────────
    for part_num, chunk in enumerate(chunks, start=1):
        op   = xml_path(part_num, total_parts)
        root = ET.Element(crs("CRS_OECD"), {"version": "3.0"})

        build_message_spec(root, header_row, tc, part_num, total_parts)

        crs_body = sub(root, crs("CrsBody"))
        build_reporting_fi(crs_body, fi_row, tc)
        rg_el = sub(crs_body, crs("ReportingGroup"))

        for _, row in chunk.iterrows():
            rd       = row_dict(row)
            sheet    = rd.pop("__sheet__",    None)
            orig_idx = int(rd.pop("__orig_idx__", -1))
            acc_num  = safe(rd.get("accountnumber"))

            if not acc_num:
                continue

            if sheet == "Individual":
                build_individual_account(rg_el, rd, cp_df, pmt_df, tc)
                ind_map[orig_idx] = op
            elif sheet == "Organisation":
                build_organisation_account(rg_el, rd, cp_df, pmt_df, tc)
                org_map[orig_idx] = op

            # ControllingPerson rows travel with their parent account
            for ci, cp_row in cp_df.iterrows():
                if safe(str(cp_row.get("accountnumber"))) == acc_num:
                    cp_map[ci] = op

            # Payment rows travel with their parent account
            for pi, pmt_row in pmt_df.iterrows():
                if safe(str(pmt_row.get("accountnumber"))) == acc_num:
                    pmt_map[pi] = op

        write_xml(root, op)
        print(f"  Part {part_num:02d}/{total_parts:02d} "
              f"— {len(chunk):>4} account row(s) → {Path(op).name}")
        xml_files.append(op)

    # ── Write Excel with filepath column ───────────────────────────────────────
    write_filepath_column(input_path, excel_out,
                          ind_map, org_map, cp_map, pmt_map)

    print(f"\n  📊  Excel with filepath → {Path(excel_out).name}")
    print(f"\n{'='*60}")
    print(f"  Done. {total_parts} XML file(s) + 1 Excel file generated.")
    print(f"{'='*60}\n")

    return xml_files, excel_out

# ── CLI ────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="CRS Excel to XML Converter v3.0",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("input",
                        help="Input Excel file (.xlsx)")
    parser.add_argument("--split", "-s", type=int, default=1,
                        metavar="N",
                        help="Number of XML files to split into (default: 1)")
    parser.add_argument("--out", "-o", default=None,
                        metavar="PREFIX",
                        help="Output filename prefix (default: input filename stem)")
    args = parser.parse_args()

    if args.split < 1:
        print("Error: --split must be >= 1")
        sys.exit(1)

    convert(args.input, n_splits=args.split, output_prefix=args.out)
