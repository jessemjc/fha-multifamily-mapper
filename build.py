"""
Rebuilds index.html from three live HUD spreadsheets.

Run with: python build.py
Requires: pandas, openpyxl, zipcodes, requests

This script:
 1. Downloads three HUD spreadsheets fresh from their stable URLs
 2. Cross-references them by FHA project number / property ID (confirms
    each property is an active, confirmed multifamily project)
 3. Attaches a zip-code centroid lat/lon/county to each property (offline lookup)
 4. Classifies each property's business type, financing bucket, and pulls
    subsidy/Section 8/use-restriction/tax-credit/tax-exempt-bond flags
 5. Injects the resulting data into template.html to produce index.html

Note on the third file: its own data.gov catalog listing shows "last
updated 2020," but the file's actual content includes dates well past
that (e.g. 2023 occupancy dates), so the catalog metadata appears to
just be stale rather than the underlying file being frozen. Worth an
occasional manual sanity-check if the subsidy/tax-credit numbers ever
look suspiciously unchanged month over month.
"""
import pandas as pd
import zipcodes
import json
import re
import requests

PROPERTY_ADDRESSES_URL = "https://www.hud.gov/sites/dfiles/Housing/documents/InsuredActiveMultifamilyFHAPropertyAddresses.xlsx"
ACTIVE_MORTGAGES_URL = "https://www.hud.gov/sites/default/files/Housing/documents/FHA-BF90-RM-A.xlsx"
PORTFOLIO_DATA_URL = "https://www.hud.gov/sites/dfiles/Housing/documents/activeportfoliopropdata.xlsx"

def download(url, dest):
    print(f"Downloading {url} ...")
    resp = requests.get(url, timeout=120, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    with open(dest, "wb") as f:
        f.write(resp.content)
    print(f"  -> saved {dest} ({len(resp.content)/1024/1024:.1f} MB)")

download(PROPERTY_ADDRESSES_URL, "property_addresses.xlsx")
download(ACTIVE_MORTGAGES_URL, "active_mortgages.xlsx")
download(PORTFOLIO_DATA_URL, "portfolio_data.xlsx")

# The Firm Commitments file's URL has the fiscal quarter baked into its
# filename (e.g. "...FY26-Q3.xlsx"), so it changes every quarter. Rather
# than hardcode a URL that will go stale, find the current one by
# scraping the stable landing page for a link matching the pattern.
FIRM_COMMITMENTS_LANDING_PAGE = "https://www.hud.gov/hud-partners/multifamily-data"
firm_commitments_path = None
try:
    resp = requests.get(FIRM_COMMITMENTS_LANDING_PAGE, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    m = re.search(
        r'https://www\.hud\.gov/sites/default/files/Housing/documents/'
        r'FHA-MF-Firm-Commitments-and-Endorsements-Database-FY\d+-FY\d+-Q\d+\.xlsx',
        resp.text
    )
    if m:
        download(m.group(0), "firm_commitments.xlsx")
        firm_commitments_path = "firm_commitments.xlsx"
    else:
        print("Could not find a Firm Commitments file link on the landing page — skipping pipeline deals this run.")
except Exception as e:
    print(f"Could not fetch/download Firm Commitments file ({e}) — skipping pipeline deals this run.")

# --- Pull the "as of" freshness date HUD publishes on each source page ---
# (these are the dates HUD itself lists, not a technical file-modified
# timestamp — falls back gracefully to "unknown" if the page format
# ever changes and the regex stops matching).
def extract_hud_date(html, filename_hint):
    idx = html.find(filename_hint)
    if idx == -1:
        return None
    window = html[idx: idx + 400]
    m = re.search(r'as of\s+(\d{1,2}/\d{1,2}/\d{4})', window, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r'Current as of\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})', window, re.IGNORECASE)
    if m:
        return m.group(1)
    return None

def fetch_hud_date(page_url, filename_hint):
    try:
        resp = requests.get(page_url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        return extract_hud_date(resp.text, filename_hint)
    except Exception as e:
        print(f"Could not fetch HUD date from {page_url}: {e}")
        return None

from datetime import datetime, timezone
mortgage_date = fetch_hud_date("https://www.hud.gov/hud-partners/multifamily-fhasl-active", "FHA-BF90-RM-A.xlsx")
addresses_date = fetch_hud_date("https://www.hud.gov/hud-partners/multifamily-preservation", "InsuredActiveMultifamilyFHAPropertyAddresses.xlsx")
portfolio_date = fetch_hud_date("https://www.hud.gov/hud-partners/multifamily-preservation", "activeportfoliopropdata.xlsx")
build_date = datetime.now(timezone.utc).strftime("%B %d, %Y")

# The Firm Commitments file states its own "Data run" date directly inside
# the spreadsheet (row 4), so pull it from there instead of a webpage.
firm_commitments_date = None
if firm_commitments_path:
    try:
        fc_meta = pd.read_excel(firm_commitments_path, sheet_name="Firm Commitments", header=None, nrows=4)
        m = re.search(r'Data run:\s*([\d/]+)', str(fc_meta.iloc[3, 0]))
        if m:
            firm_commitments_date = m.group(1)
    except Exception as e:
        print(f"Could not read Firm Commitments data-run date: {e}")

print("HUD dates found -> mortgage:", mortgage_date, "| addresses:", addresses_date,
      "| portfolio:", portfolio_date, "| firm commitments:", firm_commitments_date)

f1 = "property_addresses.xlsx"
f2 = "active_mortgages.xlsx"
f3 = "portfolio_data.xlsx"

df1 = pd.read_excel(f1, sheet_name="Sheet1")
df2 = pd.read_excel(f2, sheet_name="sheet1", header=1)

df1['fha_number'] = df1['fha_number'].astype(str).str.strip().str.zfill(8)
df2['proj_num'] = df2['HUD PROJECT NUMBER'].astype(str).str.strip().str.zfill(8)

# Only keep df2 columns we need for enrichment (avoid dup city/state/zip - use df1's address data as primary)
df2_small = df2[['proj_num','UNITS','INITIAL ENDORSEMENT DATE','FINAL ENDORSEMENT DATE',
                  'ORIGINAL MORTGAGE AMOUNT','AMORITIZED PRINCIPAL BALANCE','HOLDER NAME','HOLDER CITY','HOLDER STATE',
                  'SECTION OF ACT CODE','SOA CATEGORY/SUB CATEGORY','BUSINESS_TYPE',
                  'TC','TE']].copy()
BT_CODE = {'MF Residential':'R', 'MF Healthcare':'H', 'MF Hospitals':'P'}

# --- Third data source: Active Portfolio Property Data (subsidy/restriction flags) ---
sheet1 = pd.read_excel(f3, sheet_name='Step_01_Property_Level_data')
sheet2 = pd.read_excel(f3, sheet_name='All active Properties with FHA ')
sheet2['fha_number'] = sheet2['fha_number'].astype(str).str.strip().str.zfill(8)
portfolio = sheet1.merge(sheet2[['property_id','fha_number']], on='property_id', how='inner')
portfolio = portfolio[portfolio['is_insured_ind'] == 'Y']
portfolio = portfolio[['fha_number','is_subsidized_ind','is_sec8_ind','has_use_restriction_ind']].drop_duplicates(subset='fha_number')

# --- Financing-type bucket scheme (residential properties only) ---
BUCKET_OF = {}
LABEL_OF = {}

def _assign(codes, bucket, label_fn):
    for c in codes:
        BUCKET_OF[c] = bucket
        LABEL_OF[c] = label_fn(c)

_assign(['221(d)(4)MKT', '221(d)(4)/244'], '221d4', lambda c: '221(d)(4)')
_assign(['207/223(f)', '207/223(f)/223(e)'], '223f', lambda c: '223(f)')
_assign(['223(a)(7)/221(d)(4)M', '223a7/221d4/223e', '223(a)(7)/221(d)(4)/'],
        '223a7_221d4', lambda c: '223(a)(7) refi of 221(d)(4)')
_assign(['223(a)(7)/207/223(f)', '223(a)(7)/207'],
        '223a7_223f', lambda c: '223(a)(7) refi of 223(f)')
_assign(['542(c)'], 'risk_share', lambda c: '542(c) HFA Risk Sharing')
_assign(['542(b)'], 'risk_share', lambda c: '542(b) QPE Risk Sharing')

OTHER_LABELS = {
    '(unspecified)': 'Unspecified',
    '207': '207 Mobile Home Park',
    '213': '213 Cooperative Housing (New Construction/Rehab)',
    '213(i)': '213(i) Consumer Cooperative',
    '220': '220 Urban Renewal Housing',
    '221(d)(3)MKT': '221(d)(3) Market Rate',
    '223(a)(7)/213': '223(a)(7) refi of 213 Co-op',
    '223(a)(7)/220': '223(a)(7) refi of 220 Urban Renewal',
    '223(a)(7)/220/223(e)': '223(a)(7) refi of 220 (declining area)',
    '223(a)(7)/221': '223(a)(7) refi of 221(d)(3)/(d)(4)',
    '223(a)(7)/221(d)(3)': '223(a)(7) refi of 221(d)(3)',
    '223(a)(7)/221(d)(3)M': '223(a)(7) refi of 221(d)(3) Market Rate',
    '223(a)(7)/231': '223(a)(7) refi of 231 Elderly Housing',
    '223(a)(7)/236(j)(1)': '223(a)(7) refi of 236',
    '223(a)(7)/241': '223(a)(7) refi of 241(a) improvement loan',
    '223a7/221d3/223e': '223(a)(7) refi of 221(d)(3) (declining area)',
    '223a7/241f/236': '223(a)(7) refi of 241(f) equity loan on 236',
    '231': '231 Elderly Housing',
    '241(f)/221BMIR': '241(f) equity loan on 221 BMIR',
    '241/213': '241(a) improvement loan on 213 Co-op',
    '241/220': '241(a) improvement loan on 220 Urban Renewal',
    '241/221BMIR': '241(a) improvement loan on 221 BMIR',
    '241/221MIR': '241(a) improvement loan on 221(d) Market Rate',
    '241/223(f)': '241(a) improvement loan on 223(f) apartments',
    '241/236': '241(a) improvement loan on 236',
}
for c, lbl in OTHER_LABELS.items():
    BUCKET_OF[c] = 'other'
    LABEL_OF[c] = lbl
# dedupe proj_num in df2 (should already be unique)
df2_small = df2_small.drop_duplicates(subset='proj_num')

merged = df1.merge(df2_small, left_on='fha_number', right_on='proj_num', how='inner')
merged = merged.merge(portfolio, on='fha_number', how='left')
print("Merged rows (confirmed active MF projects):", len(merged))
print("Rows with portfolio subsidy/restriction data matched:", merged['is_subsidized_ind'].notna().sum())

# build zip -> (lat, lon, county) lookup
zip_lookup = {}
for z in zipcodes.list_all():
    lat = float(z['lat']) if z['lat'] else None
    lon = float(z['long']) if z['long'] else None
    county = z['county'] if z['county'] else None
    zip_lookup[z['zip_code']] = (lat, lon, county)

def clean_zip(z):
    s = str(z).strip()
    s = re.sub(r'\D', '', s)
    return s.zfill(5)[:5] if s else ''

records = []
no_geo = 0
for _, row in merged.iterrows():
    zip5 = clean_zip(row['zip_code'])
    lat, lon, county = zip_lookup.get(zip5, (None, None, None))
    if lat is None:
        no_geo += 1
    addr2 = str(row['address_line2_text']).strip() if pd.notna(row['address_line2_text']) else ''
    full_addr = str(row['address_line1_text']).strip()
    if addr2 and addr2.lower() != 'nan':
        full_addr += f", {addr2}"

    endorse = row['INITIAL ENDORSEMENT DATE']
    endorse_str = ''
    if pd.notna(endorse):
        try:
            endorse_str = pd.to_datetime(endorse).strftime('%Y-%m-%d')
        except Exception:
            endorse_str = str(endorse)

    final_endorse = row['FINAL ENDORSEMENT DATE']
    final_endorse_str = ''
    if pd.notna(final_endorse):
        try:
            final_endorse_str = pd.to_datetime(final_endorse).strftime('%Y-%m-%d')
        except Exception:
            final_endorse_str = str(final_endorse)

    mortgage_amt = row['ORIGINAL MORTGAGE AMOUNT']
    mortgage_amt = float(mortgage_amt) if pd.notna(mortgage_amt) else None

    amortized_balance = row['AMORITIZED PRINCIPAL BALANCE']
    amortized_balance = float(amortized_balance) if pd.notna(amortized_balance) else None

    units = row['UNITS']
    units = int(units) if pd.notna(units) else None

    def yn(val):
        # Portfolio flags: 'Y'/'N'/NaN (no match) -> True/False/None
        if pd.isna(val):
            return None
        return str(val).strip().upper() == 'Y'

    rec = {
        "n": str(row['property_name_text']).strip(),
        "a": full_addr,
        "c": str(row['city_name_text']).strip(),
        "s": str(row['state_code']).strip(),
        "z": zip5,
        "county": county or '',
        "f": row['fha_number'],
        "u": units,
        "e": endorse_str,
        "fe": final_endorse_str,
        "m": round(mortgage_amt) if mortgage_amt is not None else None,
        "bal": round(amortized_balance) if amortized_balance is not None else None,
        "h": str(row['HOLDER NAME']).strip() if pd.notna(row['HOLDER NAME']) else '',
        "so": str(row['soa_numeric_name']).strip() if pd.notna(row['soa_numeric_name']) and str(row['soa_numeric_name']).strip() else '(unspecified)',
        "la": round(lat, 4) if lat is not None else None,
        "lo": round(lon, 4) if lon is not None else None,
        "bt": BT_CODE.get(str(row['BUSINESS_TYPE']).strip(), 'R') if pd.notna(row['BUSINESS_TYPE']) else 'R',
        "sub": yn(row.get('is_subsidized_ind')),
        "sec8": yn(row.get('is_sec8_ind')),
        "restr": yn(row.get('has_use_restriction_ind')),
        "tc": bool(pd.notna(row.get('TC'))),
        "te": bool(pd.notna(row.get('TE'))),
    }
    if rec['bt'] == 'R':
        so_key = rec['so']
        rec['grp'] = BUCKET_OF.get(so_key, 'other')
        rec['fl'] = LABEL_OF.get(so_key, so_key)
    else:
        rec['grp'] = ''
        rec['fl'] = rec['so']
    records.append(rec)

print("Records without geocoding (bad/missing zip):", no_geo, "of", len(records))

# ============================================================
# FIRM COMMITMENTS / PIPELINE DEALS
# Adds two categories of properties not yet in the main dataset:
#  - "catching_up": already Initially Endorsed (i.e. actually insured),
#    just not yet reflected in this month's Active Mortgages snapshot
#  - "pipeline": Firm Commitment issued/amended/reissued/reopened but
#    not yet closed/insured at all — filtered to the last 12 months so
#    old, likely-dead commitments that never formally closed don't show
#    up as if they're live opportunities.
# Both categories only have City/State (no street address or zip) from
# HUD's source file, so location is a city (or, rarely, state) centroid
# computed entirely offline from the same zip database used everywhere
# else — never a live geocoding call.
# ============================================================
from collections import defaultdict as _defaultdict, Counter as _Counter
import difflib as _difflib

FACILITY_TO_BT = {
    'Apartment': 'R', 'Coop': 'R', 'Mobile Home Park': 'R', 'Single-Room Occupancy': 'R',
    "Asst'd Living": 'H', 'Nursing Home/Intermediate Care Facility': 'H', 'Board and Care': 'H',
    'Hospital': 'P',
}

if firm_commitments_path:
    try:
        city_lat = _defaultdict(list)
        city_lon = _defaultdict(list)
        city_county = _defaultdict(_Counter)
        state_lat = _defaultdict(list)
        state_lon = _defaultdict(list)
        for z in zipcodes.list_all():
            if not z['lat'] or not z['long']:
                continue
            lat_f, lon_f = float(z['lat']), float(z['long'])
            key = (z['city'].upper(), z['state'])
            city_lat[key].append(lat_f)
            city_lon[key].append(lon_f)
            if z['county']:
                city_county[key][z['county']] += 1
            state_lat[z['state']].append(lat_f)
            state_lon[z['state']].append(lon_f)

        city_centroids = {k: (sum(v)/len(v), sum(city_lon[k])/len(city_lon[k])) for k, v in city_lat.items()}
        state_centroids = {k: (sum(v)/len(v), sum(state_lon[k])/len(state_lon[k])) for k, v in state_lat.items()}
        cities_by_state = _defaultdict(list)
        for (city, state) in city_centroids:
            cities_by_state[state].append(city)

        def find_location(city_raw, state_raw):
            city = str(city_raw).strip().upper()
            state = str(state_raw).strip().upper()
            if (city, state) in city_centroids:
                lat, lon = city_centroids[(city, state)]
                county = city_county[(city, state)].most_common(1)[0][0] if city_county[(city, state)] else ''
                return lat, lon, county
            variants = {city.replace('.', ''), city.replace('ST.', 'SAINT'), city.replace('ST ', 'SAINT '),
                        city.replace('.', '').replace('ST ', 'SAINT ')}
            for v in variants:
                if (v, state) in city_centroids:
                    lat, lon = city_centroids[(v, state)]
                    county = city_county[(v, state)].most_common(1)[0][0] if city_county[(v, state)] else ''
                    return lat, lon, county
            candidates = cities_by_state.get(state, [])
            close = _difflib.get_close_matches(city, candidates, n=1, cutoff=0.8) if candidates else []
            if close:
                lat, lon = city_centroids[(close[0], state)]
                county = city_county[(close[0], state)].most_common(1)[0][0] if city_county[(close[0], state)] else ''
                return lat, lon, county
            if state in state_centroids:
                lat, lon = state_centroids[state]
                return lat, lon, ''
            return None, None, ''

        fc = pd.read_excel(firm_commitments_path, sheet_name="Firm Commitments", header=8)
        fc['fha8'] = fc['FHA Number'].astype(str).str.strip().str.zfill(8)
        existing_fha = set(r['f'] for r in records)

        pipeline_statuses = {'Firm Commitment Issued', 'Firm Commitment Amended',
                              'Firm Commitment Reissued', 'Firm Commitment Reopened (Previously Expired)'}

        pending_records = []
        for _, row in fc.iterrows():
            fha8 = row['fha8']
            if fha8 in existing_fha:
                continue
            status = str(row['Current Status']).strip()

            if status == 'Initially Endorsed':
                pending_type = 'catching_up'
            elif status in pipeline_statuses:
                # No age cutoff here — every pipeline deal is included
                # regardless of how old its Firm Commitment date is.
                # Age-based exclusion is a frontend filter instead
                # (default: no limit), adjustable anytime with no new data.
                pending_type = 'pipeline'
            else:
                continue

            lat, lon, county = find_location(row['Project City'], row['Project State'])

            mortgage_amt = row.get('Mortgage Amount')
            mortgage_amt = float(mortgage_amt) if pd.notna(mortgage_amt) else None
            units = row.get('Total Units')
            units = int(units) if pd.notna(units) else None
            activity_date_str = ''
            if pd.notna(row.get('Firm Activity Date')):
                try:
                    activity_date_str = pd.to_datetime(row['Firm Activity Date']).strftime('%Y-%m-%d')
                except Exception:
                    pass

            facility = str(row.get('Facility Type', '')).strip()
            category = str(row.get('Program Category', '')).strip() if pd.notna(row.get('Program Category')) else ''

            pending_records.append({
                "n": str(row['Project Name']).strip(),
                "a": "",
                "c": str(row['Project City']).strip(),
                "s": str(row['Project State']).strip(),
                "z": "",
                "county": county,
                "f": fha8,
                "u": units,
                "e": activity_date_str,
                "fe": "",
                "m": round(mortgage_amt) if mortgage_amt is not None else None,
                "bal": None,
                "h": str(row.get('Lender Name for Firm Activity', '')).strip() if pd.notna(row.get('Lender Name for Firm Activity')) else '',
                "so": category,
                "la": round(lat, 4) if lat is not None else None,
                "lo": round(lon, 4) if lon is not None else None,
                "bt": FACILITY_TO_BT.get(facility, 'R'),
                "sub": None, "sec8": None, "restr": None,
                "tc": bool(pd.notna(row.get('LIHTC'))),
                "te": bool(pd.notna(row.get('Tax Exempt Bonds'))),
                "grp": "",
                "fl": category,
                "pending": pending_type,
                "pendingStatus": status,
            })

        records.extend(pending_records)
        catching_up_n = sum(1 for r in pending_records if r['pending'] == 'catching_up')
        pipeline_n = sum(1 for r in pending_records if r['pending'] == 'pipeline')
        print(f"Added {len(pending_records)} pending/pipeline properties "
              f"({catching_up_n} catching up, {pipeline_n} pipeline, all ages — age filtering is now a frontend option)")
    except Exception as e:
        print(f"Could not process Firm Commitments file ({e}) — skipping pipeline deals (main dataset unaffected).")

data_json = json.dumps(records)
print("Records:", len(records), "| Embedded data size (MB):", round(len(data_json)/1024/1024, 2))

metadata = {
    "mortgageDate": mortgage_date or "unknown",
    "addressesDate": addresses_date or "unknown",
    "portfolioDate": portfolio_date or "unknown",
    "firmCommitmentsDate": firm_commitments_date or "unknown",
    "buildDate": build_date,
}
metadata_json = json.dumps(metadata)

# Inject into the HTML template to produce the final, deployable index.html
with open("template.html", "r", encoding="utf-8") as f:
    html = f.read()

safe_json = data_json.replace("</script>", "<\\/script>")
html = html.replace("__DATA_PLACEHOLDER__", safe_json)
html = html.replace("__METADATA_PLACEHOLDER__", metadata_json.replace("</script>", "<\\/script>"))

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"Wrote index.html ({len(html)/1024/1024:.2f} MB)")
