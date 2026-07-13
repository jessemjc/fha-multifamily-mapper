"""
Rebuilds index.html from the two live HUD spreadsheets.

Run with: python build.py
Requires: pandas, openpyxl, zipcodes, requests

This script:
 1. Downloads the two HUD spreadsheets fresh from their stable URLs
 2. Cross-references them by FHA project number (confirms each property
    is an active, confirmed multifamily project)
 3. Attaches a zip-code centroid lat/lon to each property (offline lookup)
 4. Classifies each property's business type and financing bucket
 5. Injects the resulting data into template.html to produce index.html
"""
import pandas as pd
import zipcodes
import json
import re
import requests

ACTIVE_MORTGAGES_URL = "https://www.hud.gov/sites/default/files/Housing/documents/FHA-BF90-RM-A.xlsx"
PROPERTY_ADDRESSES_URL = "https://www.hud.gov/sites/dfiles/Housing/documents/InsuredActiveMultifamilyFHAPropertyAddresses.xlsx"

def download(url, dest):
    print(f"Downloading {url} ...")
    resp = requests.get(url, timeout=120, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    with open(dest, "wb") as f:
        f.write(resp.content)
    print(f"  -> saved {dest} ({len(resp.content)/1024/1024:.1f} MB)")

download(PROPERTY_ADDRESSES_URL, "property_addresses.xlsx")
download(ACTIVE_MORTGAGES_URL, "active_mortgages.xlsx")

f1 = "property_addresses.xlsx"
f2 = "active_mortgages.xlsx"

df1 = pd.read_excel(f1, sheet_name="Sheet1")
df2 = pd.read_excel(f2, sheet_name="sheet1", header=1)

df1['fha_number'] = df1['fha_number'].astype(str).str.strip().str.zfill(8)
df2['proj_num'] = df2['HUD PROJECT NUMBER'].astype(str).str.strip().str.zfill(8)

# Only keep df2 columns we need for enrichment (avoid dup city/state/zip - use df1's address data as primary)
df2_small = df2[['proj_num','UNITS','INITIAL ENDORSEMENT DATE','FINAL ENDORSEMENT DATE',
                  'ORIGINAL MORTGAGE AMOUNT','HOLDER NAME','HOLDER CITY','HOLDER STATE',
                  'SECTION OF ACT CODE','SOA CATEGORY/SUB CATEGORY','BUSINESS_TYPE']].copy()
BT_CODE = {'MF Residential':'R', 'MF Healthcare':'H', 'MF Hospitals':'P'}

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
print("Merged rows (confirmed active MF projects):", len(merged))

# build zip -> lat/lon lookup
zip_lookup = {}
for z in zipcodes.list_all():
    zip_lookup[z['zip_code']] = (float(z['lat']), float(z['long'])) if z['lat'] and z['long'] else (None, None)

def clean_zip(z):
    s = str(z).strip()
    s = re.sub(r'\D', '', s)
    return s.zfill(5)[:5] if s else ''

records = []
no_geo = 0
for _, row in merged.iterrows():
    zip5 = clean_zip(row['zip_code'])
    lat, lon = zip_lookup.get(zip5, (None, None))
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

    mortgage_amt = row['ORIGINAL MORTGAGE AMOUNT']
    mortgage_amt = float(mortgage_amt) if pd.notna(mortgage_amt) else None

    units = row['UNITS']
    units = int(units) if pd.notna(units) else None

    rec = {
        "n": str(row['property_name_text']).strip(),
        "a": full_addr,
        "c": str(row['city_name_text']).strip(),
        "s": str(row['state_code']).strip(),
        "z": zip5,
        "f": row['fha_number'],
        "u": units,
        "e": endorse_str,
        "m": round(mortgage_amt) if mortgage_amt is not None else None,
        "h": str(row['HOLDER NAME']).strip() if pd.notna(row['HOLDER NAME']) else '',
        "so": str(row['soa_numeric_name']).strip() if pd.notna(row['soa_numeric_name']) and str(row['soa_numeric_name']).strip() else '(unspecified)',
        "la": round(lat, 4) if lat is not None else None,
        "lo": round(lon, 4) if lon is not None else None,
        "bt": BT_CODE.get(str(row['BUSINESS_TYPE']).strip(), 'R') if pd.notna(row['BUSINESS_TYPE']) else 'R',
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

data_json = json.dumps(records)
print("Records:", len(records), "| Embedded data size (MB):", round(len(data_json)/1024/1024, 2))

# Inject into the HTML template to produce the final, deployable index.html
with open("template.html", "r", encoding="utf-8") as f:
    html = f.read()

safe_json = data_json.replace("</script>", "<\\/script>")
html = html.replace("__DATA_PLACEHOLDER__", safe_json)

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"Wrote index.html ({len(html)/1024/1024:.2f} MB)")
