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
    'MEUI': 'meui',
    'Calculated MEUI': 'meui_calc',
    'TEDI': 'tedi',
    "Calc'd TEDI": 'tedi_calc',
    'Rounded ACH': 'ach',
    'GHGs': 'ghg',
    'CEILINS': 'ceil_rsi',
    'MAINWALLINS': 'wall_rsi',
    'FNDWALLINS': 'fndwall_rsi',
    'HEATEDFLOORAREA': 'floor_area',
}

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
    'TYPEOFHOUSE': 'house_type',
}

CITY_FIELD = 'CLIENTCITY'
YEAR_FIELD = 'YEARBUILT'

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
    return aliases.get(fixed, fixed)


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


def to_float(raw):
    if raw is None or raw == '':
        return None
    try:
        return float(raw)
    except ValueError:
        return None


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

    all_names = [CITY_FIELD, YEAR_FIELD] + list(NUMERIC_FIELDS) + list(CATEGORY_FIELDS)
    missing = [n for n in all_names if n not in name_to_letter]
    if missing:
        raise SystemExit(f'Columns not found in header: {missing}')
    wanted_letters = {name_to_letter[n] for n in all_names}
    letter_to_name = {name_to_letter[n]: n for n in all_names}

    cube = defaultdict(lambda: {
        'count': 0,
        'num': defaultdict(lambda: [0.0, 0]),  # [sum, count]
        'cat': defaultdict(lambda: defaultdict(int)),
    })
    city_counts = defaultdict(int)
    region_counts = defaultdict(int)
    year_counts = defaultdict(int)
    excluded_unmapped = defaultdict(int)
    # key -> lowercased value -> {original casing: count}, so values that only
    # differ by case (e.g. "Natural gas" vs "Natural Gas") can be merged into
    # whichever casing is most common, without mangling mixed-case values
    # like "Furnace w/o AC".
    case_variants = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    total_rows = 0

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
            year = normalize_year(rowvals.get(YEAR_FIELD))
            year_key = year if year is not None else 'Unknown'
            region = CITY_TO_REGION.get(city, UNMAPPED_LABEL)

            if region == UNMAPPED_LABEL:
                excluded_unmapped[city] += 1
                continue
            total_rows += 1

            city_counts[city] += 1
            region_counts[region] += 1
            year_counts[year_key] += 1

            cell = cube[(city, year_key)]
            cell['count'] += 1
            cell['region'] = region
            for field, key in NUMERIC_FIELDS.items():
                fv = to_float(rowvals.get(field))
                if fv is not None:
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

            if total_rows % 20000 == 0:
                print(f'  ... {total_rows} rows')

    print(f'Done: {total_rows} rows, {len(cube)} (city, year) cells')

    # Merge category values that only differ by case, keeping the most
    # common casing for each.
    canonical = {
        key: {lower: max(variants.items(), key=lambda kv: kv[1])[0] for lower, variants in by_lower.items()}
        for key, by_lower in case_variants.items()
    }
    merged_count = 0
    for cell in cube.values():
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

    cells_out = []
    for (city, year_key), cell in cube.items():
        cells_out.append({
            'city': city,
            'region': cell['region'],
            'year': year_key,
            'count': cell['count'],
            'num': {k: v for k, v in cell['num'].items()},
            'cat': {k: dict(v) for k, v in cell['cat'].items()},
        })

    out = {
        'generated_at': datetime.datetime.now().isoformat(timespec='seconds'),
        'source_file': wb_path.name,
        'total_rows': total_rows,
        'cities': sorted(city_counts, key=lambda c: -city_counts[c]),
        'city_counts': city_counts,
        'regions': sorted(region_counts, key=lambda r: -region_counts[r]),
        'region_counts': region_counts,
        'years': sorted([y for y in year_counts if isinstance(y, int)]),
        'cells': cells_out,
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
