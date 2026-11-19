"""
Rebuilds dashboard/data.json from the "Conditioned Data" tab of
"NRCan Database Analysis.xlsx" (the sibling folder's main workbook).

Re-run this any time the workbook has new data, then re-publish the
dashboard artifact with the refreshed data.json.

Usage:
    python extract_data.py ["path to workbook.xlsx"]

Pure standard library (no pandas/openpyxl needed) - streams the sheet XML
directly so it can handle a 140k+ row sheet without loading the whole
workbook into memory.
"""
import datetime
import json
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

from regions import CITY_TO_REGION, UNMAPPED_LABEL

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
SHEET_NAME = 'Conditioned Data'

DEFAULT_WORKBOOK = Path(__file__).parent.parent / 'NRCan Database Analysis.xlsx'
OUT_PATH = Path(__file__).parent / 'data.json'

NUMERIC_FIELDS = {
    # Use the calculated MEUI/TEDI columns, not the raw "MEUI"/"TEDI" ones:
    # the raw columns use 0 as a "not computed" sentinel for ~66% of rows
    # (physically implausible for a real home) while the calculated columns
    # are consistently populated.
    'Calculated MEUI': 'meui',
    "Calc'd TEDI": 'tedi',
    'Rounded ACH': 'ach',
    'GHGs': 'ghg',
    'CEILINS': 'ceil_rsi',
    'MAINWALLINS': 'wall_rsi',
    'FNDWALLINS': 'fndwall_rsi',
    'HEATEDFLOORAREA': 'floor_area',
    'NLA': 'nla',
    'NLR': 'nlr',
    'EGHDESHTLOSS': 'design_heat_loss',
    'ERSDesCoolLoss': 'design_cool_load',
}

# These columns use 0 as a "not computed / not applicable" sentinel instead
# of leaving the cell blank. Treat 0 as missing for these keys so it doesn't
# drag averages toward zero.
ZERO_AS_MISSING = {'nlr', 'design_cool_load'}

CATEGORY_FIELDS = {
    'Step Achieved': 'step_achieved',
    'MEUI Step Achieved': 'meui_step',
    'TEDI Step Achieved': 'tedi_step',
    'ACH Step Achieved': 'ach_step',
    'EL Achieved': 'el_achieved',
    'Has heat pump?': 'heat_pump',
    'has cooling?': 'cooling',
    'Primary Heating System': 'heating_system',
    'Primary Space Heating Fuel': 'heating_fuel',
    'PDHWTYPE': 'dhw_type',
    'PDHWFUEL': 'dhw_fuel',
    'AIRCONDTYPE': 'cooling_type',
    'BUILDINGTYPE': 'building_type',
}

CITY_FIELD = 'CLIENTCITY'
YEAR_FIELD = 'YEARBUILT'
# An entry can be assigned to a year two ways: when the home was built, or
# when it was entered/evaluated. The dashboard lets the user pick, so both
# are extracted into their own cube (see YEAR_MODES below).
ENTRY_DATE_FIELD = 'ENTRYDATE'
# House type (single detached, row house, duplex, ...) and MURB/non-MURB
# split the cube like city/year/region rather than being just another
# category histogram, so those filters can actually exclude homes, not just
# relabel them.
HOUSE_TYPE_FIELD = 'TYPEOFHOUSE'
MURB_FIELD = 'UNITSMURBS'

MIN_PLAUSIBLE_YEAR = 1990
MAX_PLAUSIBLE_YEAR = datetime.date.today().year + 1


def col_letters(ref):
    return ''.join(ch for ch in ref if ch.isalpha())


def normalize_city(raw):
    if not raw:
        return 'Unknown'
    name = ' '.join(raw.strip().split())
    # Title-case but keep common connector words lowercase, and fix all-caps input.
    words = name.split(' ')
    small = {'of'}
    out = []
    for w in words:
        lw = w.lower()
        if lw in small:
            out.append(lw)
        else:
            out.append(lw[:1].upper() + lw[1:] if lw else w)
    fixed = ' '.join(out)
    aliases = {
        'N Vancouver': 'North Vancouver',
        'N. Vancouver': 'North Vancouver',
        'W Vancouver': 'West Vancouver',
        'W. Vancouver': 'West Vancouver',
    }
    fixed = aliases.get(fixed, fixed)
    return CITY_SPELLING_FIXES.get(fixed, fixed)


# Misspellings/typos of a real municipality, found by fuzzy-matching every
# distinct CLIENTCITY value against the others (see dashboard dev notes).
# Keyed by the value AFTER the title-casing above, mapped to the correct
# spelling. Deliberately excludes near-matches that are actually different
# real places (e.g. Invermere/Windermere, the "N Mile House" series,
# Koocanusa's distinct sub-localities).
CITY_SPELLING_FIXES = {
    # North Saanich (6 misspelled variants, per user report)
    'North Saaanich': 'North Saanich',
    'North Sasanich': 'North Saanich',
    'North Saaainch': 'North Saanich',
    'N. Saanich': 'North Saanich',
    'N.saanich': 'North Saanich',
    'North Sannich': 'North Saanich',
    # Saanich
    'Saainch': 'Saanich',
    'Sannich': 'Saanich',
    # Other municipalities
    'Fairmont Hot Springs,': 'Fairmont Hot Springs',
    'Fairmont Hotsprings': 'Fairmont Hot Springs',
    'Saltspring Island': 'Salt Spring Island',
    'New West Minster': 'New Westminster',
    'New Westminister': 'New Westminster',
    'New Wesminster': 'New Westminster',
    'New Westmister': 'New Westminster',
    'Newwestminster': 'New Westminster',
    'New Westminst': 'New Westminster',
    'Crambrook': 'Cranbrook',
    'Cranbroook': 'Cranbrook',
    'Cranbroook/rdek': 'Cranbrook/rdek',
    'Northvancouver': 'North Vancouver',
    'N.vancouver': 'North Vancouver',
    'Chrstina Lake': 'Christina Lake',
    'Fort St John': 'Fort St. John',
    'Fort St. .john': 'Fort St. John',
    'Westvancouver': 'West Vancouver',
    'Vancouverr': 'Vancouver',
    'Ancouver': 'Vancouver',
    'Vancouer': 'Vancouver',
    'Vancovuer': 'Vancouver',
    'Vancpuver': 'Vancouver',
    'Vanocuver': 'Vancouver',
    'Vanvouver': 'Vancouver',
    'Vnacouver': 'Vancouver',
    'E Vancouver': 'Vancouver',
    'Princegeorge': 'Prince George',
    'Ft St James': 'Ft. St James',
    'Pittmeadows': 'Pitt Meadows',
    'Port Alerni': 'Port Alberni',
    'Abbottsford': 'Abbotsford',
    'Sumemerland': 'Summerland',
    'Anglemount': 'Anglemont',
    'Chiliwack': 'Chilliwack',
    'Courtenaty': 'Courtenay',
    'Courtnenay': 'Courtenay',
    'Courtneay': 'Courtenay',
    'Courtnay': 'Courtenay',
    'Courtney': 'Courtenay',
    'Couretnay': 'Courtenay',
    'Parkville': 'Parksville',
    'Permberton': 'Pemberton',
    'Pentiction': 'Penticton',
    'Tswwassen': 'Tsawwassen',
    'Invemere': 'Invermere',
    'Kamploops': 'Kamloops',
    'Lilllooet': 'Lillooet',
    'Lillooett': 'Lillooet',
    'Narramata': 'Naramata',
    'Coolwood': 'Colwood',
    'Ucluelt': 'Ucluelet',
    'Ucuelet': 'Ucluelet',
    'Koocanusa-west': 'Koocanusa West',
    'Gibson': 'Gibsons',
    'West Kelonwa': 'West Kelowna',
    'Wset Kelowna': 'West Kelowna',
    'Ooteschenia': 'Ootischenia',
    'Lanztville': 'Lantzville',
    'Latnzville': 'Lantzville',
    'Kimberkey': 'Kimberley',
    'Smitherr': 'Smithers',
    'Sqaumish': 'Squamish',
    'Squamsih': 'Squamish',
    'Whislter': 'Whistler',
    'Whsitler': 'Whistler',
    'Kelowma': 'Kelowna',
    'Langely': 'Langley',
    'Nanamio': 'Nanaimo',
    'Naniamo': 'Nanaimo',
    'Naksup': 'Nakusp',
    'Suerry': 'Surrey',
    'Clucitz Lake': 'Cluculz Lake',
    'Clucutz Lake': 'Cluculz Lake',
    'Mayook Bc/rdek': 'Mayook/rdek',
    'Proctor': 'Procter',
    'Koocanusa/redk': 'Koocanusa Rdek',
    'Koocanusa Vllge/rdek': 'Koocanusa Village Rdek',
    'Kalso': 'Kaslo',
    'Smitehsr': 'Smithers',
}


HOUSE_TYPE_FIXES = {
    'Single detached': 'Single Detached',
}


def normalize_house_type(raw):
    if not raw:
        return 'Unknown'
    name = ' '.join(raw.strip().split())
    if not name:
        return 'Unknown'
    return HOUSE_TYPE_FIXES.get(name, name)


def normalize_year(raw):
    if not raw:
        return None
    try:
        y = int(float(raw))
    except ValueError:
        return None
    if y < MIN_PLAUSIBLE_YEAR or y > MAX_PLAUSIBLE_YEAR:
        return None
    return y


# Excel stores dates as a day count from 1899-12-30 (the epoch is a day
# earlier than 1900-01-01 to reproduce Excel's fictitious Feb 29 1900).
EXCEL_EPOCH = datetime.date(1899, 12, 30)


def entry_date_to_year(raw):
    if raw is None or raw == '':
        return None
    try:
        serial = float(raw)
    except ValueError:
        return None
    try:
        d = EXCEL_EPOCH + datetime.timedelta(days=serial)
    except OverflowError:
        return None
    y = d.year
    if y < MIN_PLAUSIBLE_YEAR or y > MAX_PLAUSIBLE_YEAR:
        return None
    return y


def is_murb(raw):
    v = to_float(raw)
    return bool(v)


def to_float(raw):
    if raw is None or raw == '':
        return None
    try:
        return float(raw)
    except ValueError:
        return None


# For the "limiting parameter" breakdown: whichever of airtightness/MEUI/TEDI
# step-achieved is lowest is the one holding the home back from a higher
# overall step. Ties (two or more fields sharing the lowest value) are
# grouped into a single "Multiple limiting parameters" bucket rather than a
# combinatorial label per tie, which was too messy to read as a chart.
LIMITING_FIELDS = [
    ('ACH Step Achieved', 'Airtightness'),
    ('MEUI Step Achieved', 'MEUI'),
    ('TEDI Step Achieved', 'TEDI'),
]
MULTIPLE_LIMITING_LABEL = 'Multiple limiting parameters'


def limiting_param_label(rowvals):
    vals = []
    for field, label in LIMITING_FIELDS:
        v = to_float(rowvals.get(field))
        if v is None:
            return None
        vals.append((label, v))
    min_v = min(v for _, v in vals)
    winners = [label for label, v in vals if v == min_v]
    if len(winners) > 1:
        return MULTIPLE_LIMITING_LABEL
    return winners[0]


def main():
    wb_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_WORKBOOK
    print(f'Reading {wb_path.name} ...')
    z = zipfile.ZipFile(wb_path)

    # Map sheet name -> worksheet XML part.
    wb_xml = z.read('xl/workbook.xml').decode('utf-8', 'ignore')
    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'ignore')
    import re
    sheets = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb_xml)
    relmap = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml))
    sheet_rid = next(rid for name, rid in sheets if name == SHEET_NAME)
    sheet_part = 'xl/' + relmap[sheet_rid]

    print('Loading shared strings ...')
    ss = []
    with z.open('xl/sharedStrings.xml') as f:
        for _, elem in ET.iterparse(f, events=('end',)):
            if elem.tag == NS + 'si':
                ss.append(''.join(t.text or '' for t in elem.iter(NS + 't')))
                elem.clear()
    print(f'  {len(ss)} shared strings')

    print('Reading header row ...')
    header = {}
    with z.open(sheet_part) as f:
        for _, elem in ET.iterparse(f, events=('end',)):
            if elem.tag == NS + 'row':
                if int(elem.get('r')) == 1:
                    for c in elem.iter(NS + 'c'):
                        v = c.find(NS + 'v')
                        val = v.text if v is not None else None
                        if c.get('t') == 's' and val is not None:
                            val = ss[int(val)]
                        header[col_letters(c.get('r'))] = val
                    elem.clear()
                    break
                elem.clear()
    name_to_letter = {v: k for k, v in header.items() if v}

    all_names = [CITY_FIELD, YEAR_FIELD, ENTRY_DATE_FIELD, HOUSE_TYPE_FIELD, MURB_FIELD] + list(NUMERIC_FIELDS) + list(CATEGORY_FIELDS)
    missing = [n for n in all_names if n not in name_to_letter]
    if missing:
        raise SystemExit(f'Columns not found in header: {missing}')
    wanted_letters = {name_to_letter[n] for n in all_names}
    letter_to_name = {name_to_letter[n]: n for n in all_names}

    def new_cube():
        return defaultdict(lambda: {
            'count': 0,
            'num': defaultdict(lambda: [0.0, 0]),  # [sum, count]
            'cat': defaultdict(lambda: defaultdict(int)),
        })

    # Two independent cubes, one per way of assigning an entry to a year -
    # the user picks which to view in the dashboard.
    YEAR_MODES = ('built', 'entry')
    cubes = {mode: new_cube() for mode in YEAR_MODES}
    year_counts = {mode: defaultdict(int) for mode in YEAR_MODES}
    city_counts = defaultdict(int)
    region_counts = defaultdict(int)
    house_type_counts = defaultdict(int)
    murb_counts = defaultdict(int)
    excluded_unmapped = defaultdict(int)
    # key -> lowercased value -> {original casing: count}, so values that only
    # differ by case (e.g. "Natural gas" vs "Natural Gas") can be merged into
    # whichever casing is most common, without mangling mixed-case values
    # like "Furnace w/o AC".
    case_variants = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    total_rows = 0

    def apply_row(cell, rowvals):
        cell['count'] += 1
        for field, key in NUMERIC_FIELDS.items():
            fv = to_float(rowvals.get(field))
            if fv is not None and not (fv == 0 and key in ZERO_AS_MISSING):
                acc = cell['num'][key]
                acc[0] += fv
                acc[1] += 1
        for field, key in CATEGORY_FIELDS.items():
            cv = rowvals.get(field)
            cv = cv.strip() if isinstance(cv, str) else cv
            if cv:
                cell['cat'][key][cv] += 1
                if isinstance(cv, str):
                    case_variants[key][cv.lower()][cv] += 1
        limiting = limiting_param_label(rowvals)
        if limiting:
            cell['cat']['limiting_param'][limiting] += 1

    print('Streaming data rows (this can take a minute or two) ...')
    with z.open(sheet_part) as f:
        for _, elem in ET.iterparse(f, events=('end',)):
            if elem.tag != NS + 'row':
                continue
            r = int(elem.get('r'))
            if r == 1:
                elem.clear()
                continue
            rowvals = {}
            for c in elem.iter(NS + 'c'):
                letter = col_letters(c.get('r'))
                if letter not in wanted_letters:
                    continue
                v = c.find(NS + 'v')
                val = v.text if v is not None else None
                if c.get('t') == 's' and val is not None:
                    val = ss[int(val)]
                rowvals[letter_to_name[letter]] = val
            elem.clear()

            if not rowvals:
                continue

            city = normalize_city(rowvals.get(CITY_FIELD))
            year_built = normalize_year(rowvals.get(YEAR_FIELD))
            year_entry = entry_date_to_year(rowvals.get(ENTRY_DATE_FIELD))
            year_keys = {
                'built': year_built if year_built is not None else 'Unknown',
                'entry': year_entry if year_entry is not None else 'Unknown',
            }
            region = CITY_TO_REGION.get(city, UNMAPPED_LABEL)
            house_type = normalize_house_type(rowvals.get(HOUSE_TYPE_FIELD))
            murb = is_murb(rowvals.get(MURB_FIELD))

            if region == UNMAPPED_LABEL:
                excluded_unmapped[city] += 1
                continue
            total_rows += 1

            city_counts[city] += 1
            region_counts[region] += 1
            house_type_counts[house_type] += 1
            murb_counts['MURB' if murb else 'Non-MURB'] += 1

            for mode in YEAR_MODES:
                year_key = year_keys[mode]
                year_counts[mode][year_key] += 1
                cell = cubes[mode][(city, year_key, house_type, murb)]
                cell['region'] = region
                cell['house_type'] = house_type
                cell['is_murb'] = murb
                apply_row(cell, rowvals)

            if total_rows % 20000 == 0:
                print(f'  ... {total_rows} rows')

    print(f'Done: {total_rows} rows, ' + ', '.join(f'{len(cubes[m])} {m} cells' for m in YEAR_MODES))

    # Merge category values that only differ by case, keeping the most
    # common casing for each.
    canonical = {
        key: {lower: max(variants.items(), key=lambda kv: kv[1])[0] for lower, variants in by_lower.items()}
        for key, by_lower in case_variants.items()
    }
    merged_count = 0
    for mode in YEAR_MODES:
        for cell in cubes[mode].values():
            for key, hist in cell['cat'].items():
                fixed = defaultdict(int)
                for cv, n in hist.items():
                    canon = canonical.get(key, {}).get(cv.lower(), cv) if isinstance(cv, str) else cv
                    if canon != cv:
                        merged_count += 1
                    fixed[canon] += n
                cell['cat'][key] = fixed
    if merged_count:
        print(f'Merged {merged_count} case-variant category values (e.g. "Natural gas" / "Natural Gas") '
              f'into their most common casing')

    year_modes_out = {}
    for mode in YEAR_MODES:
        cells_out = []
        for (city, year_key, house_type, murb), cell in cubes[mode].items():
            cells_out.append({
                'city': city,
                'region': cell['region'],
                'year': year_key,
                'house_type': house_type,
                'is_murb': cell['is_murb'],
                'count': cell['count'],
                'num': {k: v for k, v in cell['num'].items()},
                'cat': {k: dict(v) for k, v in cell['cat'].items()},
            })
        year_modes_out[mode] = {
            'years': sorted([y for y in year_counts[mode] if isinstance(y, int)]),
            'cells': cells_out,
        }

    out = {
        'generated_at': datetime.datetime.now().isoformat(timespec='seconds'),
        'source_file': wb_path.name,
        'total_rows': total_rows,
        'cities': sorted(city_counts, key=lambda c: -city_counts[c]),
        'city_counts': city_counts,
        'regions': sorted(region_counts, key=lambda r: -region_counts[r]),
        'region_counts': region_counts,
        'house_types': sorted(house_type_counts, key=lambda h: -house_type_counts[h]),
        'house_type_counts': house_type_counts,
        'murb_counts': murb_counts,
        'year_mode_default': 'built',
        'year_modes': year_modes_out,
    }

    OUT_PATH.write_text(json.dumps(out, separators=(',', ':')), encoding='utf-8')
    size_mb = OUT_PATH.stat().st_size / 1_000_000
    print(f'Wrote {OUT_PATH} ({size_mb:.2f} MB)')

    if excluded_unmapped:
        total_excluded = sum(excluded_unmapped.values())
        print(f'\nExcluded {total_excluded} homes ({len(excluded_unmapped)} distinct city values) with no '
              f'regional district mapping in regions.py - add them there to include them:')
        for c, n in sorted(excluded_unmapped.items(), key=lambda kv: -kv[1]):
            print(f'  {n:5d}  {c}')


if __name__ == '__main__':
    main()
