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
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

from regions import CITY_TO_REGION, UNMAPPED_LABEL
from fsa import city_from_fsa

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
SHEET_NAME = 'Conditioned Data'

DEFAULT_WORKBOOK = Path(__file__).parent.parent / 'NRCan Database Analysis.xlsx'
OUT_PATH = Path(__file__).parent / 'data.json'

NUMERIC_FIELDS = {
    # Use the calculated MEUI/TEDI/NLR columns, not the raw "MEUI"/"TEDI"/
    # "NLR" ones: the raw columns use 0 as a "not computed" sentinel for a
    # large fraction of rows (physically implausible for a real home) while
    # the calculated columns are consistently populated.
    'Calculated MEUI': 'meui',
    "Calc'd TEDI": 'tedi',
    "Calc'd NLR": 'nlr',
    'Rounded ACH': 'ach',
    # GHGs is deliberately NOT here - it's paired with HEATEDFLOORAREA to
    # also derive ghg_intensity, unlike every other field in this dict, so
    # it's computed separately in apply_row()/collect_median_values()
    # instead of through this generic loop.
    'CEILINS': 'ceil_rsi',
    'MAINWALLINS': 'wall_rsi',
    'FNDWALLINS': 'fndwall_rsi',
    'HEATEDFLOORAREA': 'floor_area',
    'NLA': 'nla',
    'EGHDESHTLOSS': 'design_heat_loss',
    'ERSDesCoolLoss': 'design_cool_load',
    # Fractional (0-1) improvement over the reference house, per CSA F280 /
    # the step code's own methodology - not the per-step target columns
    # ("TEDI % Imp Step 3" etc.), which are thresholds, not achieved values.
    "Calc'd TEDI % Imp": 'tedi_imp',
    # Unlike the TEDI version, this column is already a percentage (not a
    # 0-1 fraction).
    'Calculated MEUI % Improvement': 'meui_imp',
}

# These columns use 0 as a "not computed / not applicable" sentinel instead
# of leaving the cell blank (a real home never has literally zero airtightness,
# floor area, energy use, GHGs, etc.) - treat 0 as missing for these keys so it
# doesn't drag averages/medians toward zero. Deliberately excludes tedi_imp/
# meui_imp: 0% (or negative) improvement over the reference house is a
# legitimate, real value there, not a sentinel.
ZERO_AS_MISSING = {'design_cool_load', 'design_heat_loss', 'meui', 'tedi', 'nlr', 'ach', 'nla',
                    'floor_area', 'ceil_rsi', 'wall_rsi', 'fndwall_rsi'}

CATEGORY_FIELDS = {
    'Step Achieved': 'step_achieved',
    'MEUI Step Achieved': 'meui_step',
    'TEDI Step Achieved': 'tedi_step',
    # The workbook now computes airtightness compliance via three possible
    # paths (ACH50, NLR, NLA) and combines them into this single column
    # (whichever path gives the highest step, or "FAIL" if none pass) -
    # that's the real per-home airtightness step, not the ACH50-only
    # "ACH Step Achieved" column.
    'Airtightness Step': 'ach_step',
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

# Fallback for HEATEDFLOORAREA in the per-floor-area design heat loss/gain
# calcs below, mirroring the workbook's own TEDI/MEUI fallback: used only
# when HEATEDFLOORAREA is null or 0.
FLOOR_AREA_FALLBACK_FIELD = 'FLOORAREA'

CITY_FIELD = 'CLIENTCITY'
# The municipality a row is bucketed into is primarily determined by this
# field's Forward Sortation Area (see fsa.py) - CLIENTCITY is free text (full
# of typos and inconsistent spellings) while a postal code is far more
# reliable. CITY_FIELD above is only a fallback for the rows (a large
# minority) that have no postal code at all.
PCODE_FIELD = 'CLIENTPCODE'
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
# Climate zone (4, 5, 6, 7.1, 7.2) is also a cube dimension, same reasoning.
CLIMATE_ZONE_FIELD = 'Climate Zone'
# Overall Step Achieved ('0'/'3'/'4'/'5' in this data) is also a cube
# dimension, so filtering by it can include/exclude homes from every
# aggregate, not just its own category breakdown chart.
STEP_ACHIEVED_FIELD = 'Step Achieved'

MIN_PLAUSIBLE_YEAR = 1990
MAX_PLAUSIBLE_YEAR = datetime.date.today().year + 1

# This dashboard's own median/mean toggle uses the per-cell 'median' raw
# arrays added to every cube cell above (same city/year/house-type/MURB/
# climate-zone/step grouping the dashboard already filters by, so a median
# recomputes correctly under any filter combination, just like the mean does).
#
# The sibling Victoria/Saanich as-built comparison dashboards need something
# coarser: this dashboard's own median by city+year *only* (not split further
# by house type/climate/step), so each group stays a reasonably sized set of
# homes rather than something closer to an individual record - a deliberately
# more private, separate export (not reused from the per-cell arrays above,
# which can be far more specific) for only these two cities, not the other
# ~100 in this dataset - see MEDIAN_EXPORT_KEYS below and 'median_export' in
# the output.
MEDIAN_EXPORT_CITIES = {'Victoria', 'Saanich'}
MEDIAN_EXPORT_KEYS = ['meui', 'tedi', 'ach', 'nla', 'nlr', 'floor_area', 'ghg', 'ghg_intensity',
                       'meui_imp', 'tedi_imp', 'design_heat_loss_per_area']

# The Saanich sub-dashboard additionally needs each home's metrics kept
# together (not shuffled independently per metric like median_export above),
# so it can identify its highest-GHG homes and drop those same homes from
# every other metric too, not just GHG. Scoped to Saanich alone - a step
# down in anonymity from median_export above, so it's kept to just the one
# city that actually needs it rather than extended to Victoria too.
LINKED_HOME_EXPORT_CITIES = {'Saanich'}


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
    fixed = CITY_SPELLING_FIXES.get(fixed, fixed)
    # Drop a trailing "/RDEK" or " RDEK" marker (Regional District of East
    # Kootenay) - the region is already computed separately from the city
    # name via CITY_TO_REGION, so repeating it in the municipality field is
    # redundant.
    fixed = RDEK_SUFFIX_RE.sub('', fixed).strip()
    return fixed if fixed else 'Unknown'


# Matches a trailing "/rdek" or " rdek" (with any amount of surrounding
# whitespace around the slash) at the end of a city name, case-insensitively.
RDEK_SUFFIX_RE = re.compile(r'(?:\s*/\s*|\s+)rdek\s*$', re.IGNORECASE)


# Misspellings/typos of a real municipality, found by fuzzy-matching every
# distinct CLIENTCITY value against the others (see dashboard dev notes).
# Keyed by the value AFTER the title-casing above, mapped to the correct
# spelling. Deliberately excludes near-matches that are actually different
# real places (e.g. Invermere/Windermere, the "N Mile House" series,
# Koocanusa's distinct sub-localities).
CITY_SPELLING_FIXES = {
    # RDEK (Regional District of East Kootenay) appears as a prefix, or
    # fused with another word, in a handful of entries, rather than as the
    # plain trailing "/RDEK" or " RDEK" that RDEK_SUFFIX_RE strips - map
    # these straight to the real place name.
    'Rdek/wardner': 'Wardner',
    'Rdek Wasa': 'Wasa',
    'Cbk/rdek': 'Cranbrook',
    'Ft Steele Rdek': 'Fort Steele',
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


def normalize_climate_zone(raw):
    v = to_float(raw)
    if v is None:
        return 'Unknown'
    label = str(int(v)) if v == int(v) else str(v)
    return 'Zone ' + label


def normalize_step_achieved(raw):
    if raw is None or raw == '':
        return 'Unknown'
    return raw.strip() if isinstance(raw, str) else str(raw)


# "Airtightness Step" stores "FAIL" (text) rather than 0 for homes that
# don't meet any step's threshold - normalize it to '0' so it groups with
# the other step fields' existing 0-as-Fail convention (see stepLabel() in
# the dashboard).
CATEGORY_FIELD_NORMALIZERS = {
    'Airtightness Step': lambda v: '0' if isinstance(v, str) and v.strip().upper() == 'FAIL' else v,
}


# For the "limiting parameter" breakdown: whichever of airtightness/MEUI/TEDI
# step-achieved is lowest is the one holding the home back from a higher
# overall step. Ties (two or more fields sharing the lowest value) are
# grouped into a single "Multiple limiting parameters" bucket rather than a
# combinatorial label per tie, which was too messy to read as a chart.
# Uses the combined "Airtightness Step" column (not the ACH50-only "ACH Step
# Achieved") so this matches the workbook's own overall "Step Achieved"
# formula, which is List.Min({Airtightness Step, TEDI Step Achieved,
# Overall MEUI Achieved}).
LIMITING_FIELDS = [
    ('Airtightness Step', 'Airtightness'),
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


# For the "fossil fuel burning system" breakdown: a home has one if its
# furnace, its primary water heater, or either of its supplementary
# heating/fireplace slots runs on a combustible fossil fuel. Wood-burning
# fireplaces/stoves (a common SUPPHTGFUEL value) don't count - wood isn't a
# fossil fuel.
FOSSIL_FUEL_FIELDS = ['FURNACEFUEL', 'PDHWFUEL', 'SUPPHTGFUEL1', 'SUPPHTGFUEL2']
FOSSIL_FUELS = {'natural gas', 'propane', 'oil'}
FOSSIL_FUEL_PRESENT_LABEL = 'Fossil fuel present'
FOSSIL_FUEL_ABSENT_LABEL = 'No fossil fuel'


def fossil_fuel_system_label(rowvals):
    for field in FOSSIL_FUEL_FIELDS:
        raw = rowvals.get(field)
        if isinstance(raw, str) and raw.strip().lower() in FOSSIL_FUELS:
            return FOSSIL_FUEL_PRESENT_LABEL
    return FOSSIL_FUEL_ABSENT_LABEL


def collect_median_values(rowvals):
    # Mirrors apply_row()'s own value computation (same fields, same
    # zero-as-missing/derived-ratio handling), but returns plain values
    # instead of accumulating into a [sum, count] pair.
    out = {}
    for field, key in NUMERIC_FIELDS.items():
        if key not in MEDIAN_EXPORT_KEYS:
            continue
        fv = to_float(rowvals.get(field))
        if fv is not None and not (fv == 0 and key in ZERO_AS_MISSING):
            out[key] = fv
    ghg = to_float(rowvals.get('GHGs'))
    floor_area = to_float(rowvals.get('HEATEDFLOORAREA'))
    if ghg and floor_area:
        out['ghg_intensity'] = ghg / floor_area
    if ghg:
        out['ghg'] = ghg
    area = floor_area or to_float(rowvals.get(FLOOR_AREA_FALLBACK_FIELD))
    if area:
        heat_loss = to_float(rowvals.get('EGHDESHTLOSS'))
        if heat_loss:
            out['design_heat_loss_per_area'] = heat_loss / area
    return out


def main():
    wb_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_WORKBOOK
    print(f'Reading {wb_path.name} ...')
    z = zipfile.ZipFile(wb_path)

    # Map sheet name -> worksheet XML part.
    wb_xml = z.read('xl/workbook.xml').decode('utf-8', 'ignore')
    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'ignore')
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

    all_names = [CITY_FIELD, PCODE_FIELD, 'GHGs', YEAR_FIELD, ENTRY_DATE_FIELD, HOUSE_TYPE_FIELD, MURB_FIELD, CLIMATE_ZONE_FIELD, FLOOR_AREA_FALLBACK_FIELD] + \
        list(NUMERIC_FIELDS) + list(CATEGORY_FIELDS) + ['FURNACEFUEL', 'SUPPHTGFUEL1', 'SUPPHTGFUEL2']
    missing = [n for n in all_names if n not in name_to_letter]
    if missing:
        raise SystemExit(f'Columns not found in header: {missing}')
    wanted_letters = {name_to_letter[n] for n in all_names}
    letter_to_name = {name_to_letter[n]: n for n in all_names}

    def new_cube():
        return defaultdict(lambda: {
            'count': 0,
            'num': defaultdict(lambda: [0.0, 0]),  # [sum, count]
            # Raw per-home values, same keys/grouping as 'num' above (sorted
            # before output - see below), so the dashboard can offer a median
            # alongside the mean for every chart, computed under any
            # combination of the dashboard's own filters (which already match
            # this cube's own grouping: city/year/house type/MURB/climate
            # zone/step).
            'median': defaultdict(list),
            'cat': defaultdict(lambda: defaultdict(int)),
        })

    # Two independent cubes, one per way of assigning an entry to a year -
    # the user picks which to view in the dashboard.
    YEAR_MODES = ('built', 'entry')
    cubes = {mode: new_cube() for mode in YEAR_MODES}
    # mode -> (city, year_key) -> metric key -> [values] - see MEDIAN_EXPORT_CITIES.
    median_raw = {mode: defaultdict(lambda: defaultdict(list)) for mode in YEAR_MODES}
    # mode -> (city, year_key) -> [per-home metric dicts] - see LINKED_HOME_EXPORT_CITIES.
    linked_homes_raw = {mode: defaultdict(list) for mode in YEAR_MODES}
    year_counts = {mode: defaultdict(int) for mode in YEAR_MODES}
    city_counts = defaultdict(int)
    region_counts = defaultdict(int)
    house_type_counts = defaultdict(int)
    murb_counts = defaultdict(int)
    climate_zone_counts = defaultdict(int)
    step_achieved_counts = defaultdict(int)
    excluded_unmapped = defaultdict(int)
    # key -> lowercased value -> {original casing: count}, so values that only
    # differ by case (e.g. "Natural gas" vs "Natural Gas") can be merged into
    # whichever casing is most common, without mangling mixed-case values
    # like "Furnace w/o AC".
    case_variants = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    total_rows = 0

    def add_metric(cell, key, value, collect_median=True):
        acc = cell['num'][key]
        acc[0] += value
        acc[1] += 1
        if collect_median:
            cell['median'][key].append(value)

    def apply_row(cell, rowvals):
        cell['count'] += 1
        for field, key in NUMERIC_FIELDS.items():
            fv = to_float(rowvals.get(field))
            if fv is not None and not (fv == 0 and key in ZERO_AS_MISSING):
                add_metric(cell, key, fv)
        for field, key in CATEGORY_FIELDS.items():
            cv = rowvals.get(field)
            cv = cv.strip() if isinstance(cv, str) else cv
            if field in CATEGORY_FIELD_NORMALIZERS:
                cv = CATEGORY_FIELD_NORMALIZERS[field](cv)
            if cv:
                cell['cat'][key][cv] += 1
                if isinstance(cv, str):
                    case_variants[key][cv.lower()][cv] += 1
        limiting = limiting_param_label(rowvals)
        if limiting:
            cell['cat']['limiting_param'][limiting] += 1
        cell['cat']['fossil_fuel_system'][fossil_fuel_system_label(rowvals)] += 1
        ghg = to_float(rowvals.get('GHGs'))
        floor_area = to_float(rowvals.get('HEATEDFLOORAREA'))
        if ghg and floor_area:
            add_metric(cell, 'ghg_intensity', ghg / floor_area)
        if ghg:
            add_metric(cell, 'ghg', ghg)
        area = floor_area or to_float(rowvals.get(FLOOR_AREA_FALLBACK_FIELD))
        if area:
            heat_loss = to_float(rowvals.get('EGHDESHTLOSS'))
            if heat_loss:
                add_metric(cell, 'design_heat_loss_per_area', heat_loss / area)
            cool_load = to_float(rowvals.get('ERSDesCoolLoss'))
            if cool_load:
                add_metric(cell, 'design_cool_load_per_area', cool_load / area)

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

            city = city_from_fsa(rowvals.get(PCODE_FIELD)) or normalize_city(rowvals.get(CITY_FIELD))
            year_built = normalize_year(rowvals.get(YEAR_FIELD))
            year_entry = entry_date_to_year(rowvals.get(ENTRY_DATE_FIELD))
            year_keys = {
                'built': year_built if year_built is not None else 'Unknown',
                'entry': year_entry if year_entry is not None else 'Unknown',
            }
            region = CITY_TO_REGION.get(city, UNMAPPED_LABEL)
            # "Rdek" alone (the city field just says "RDEK", with no actual
            # place name) isn't a municipality - display it as Unknown, but
            # only after the region lookup above, which still needs "Rdek"
            # to resolve to East Kootenay.
            if city == 'Rdek':
                city = 'Unknown'
            house_type = normalize_house_type(rowvals.get(HOUSE_TYPE_FIELD))
            murb = is_murb(rowvals.get(MURB_FIELD))
            climate_zone = normalize_climate_zone(rowvals.get(CLIMATE_ZONE_FIELD))
            step_achieved = normalize_step_achieved(rowvals.get(STEP_ACHIEVED_FIELD))

            if region == UNMAPPED_LABEL:
                excluded_unmapped[city] += 1
                continue
            total_rows += 1

            city_counts[city] += 1
            region_counts[region] += 1
            house_type_counts[house_type] += 1
            murb_counts['MURB' if murb else 'Non-MURB'] += 1
            climate_zone_counts[climate_zone] += 1
            step_achieved_counts[step_achieved] += 1

            median_values = collect_median_values(rowvals) if city in MEDIAN_EXPORT_CITIES else None

            for mode in YEAR_MODES:
                year_key = year_keys[mode]
                year_counts[mode][year_key] += 1
                cell = cubes[mode][(city, year_key, house_type, murb, climate_zone, step_achieved)]
                cell['region'] = region
                cell['house_type'] = house_type
                cell['is_murb'] = murb
                cell['climate_zone'] = climate_zone
                cell['step_achieved'] = step_achieved
                apply_row(cell, rowvals)
                if median_values:
                    bucket = median_raw[mode][(city, year_key)]
                    for key, v in median_values.items():
                        bucket[key].append(v)
                    if city in LINKED_HOME_EXPORT_CITIES:
                        linked_homes_raw[mode][(city, year_key)].append(median_values)

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
        for (city, year_key, house_type, murb, climate_zone, step_achieved), cell in cubes[mode].items():
            cells_out.append({
                'city': city,
                'region': cell['region'],
                'year': year_key,
                'house_type': house_type,
                'is_murb': cell['is_murb'],
                'climate_zone': cell['climate_zone'],
                'step_achieved': cell['step_achieved'],
                'count': cell['count'],
                'num': {k: v for k, v in cell['num'].items()},
                # Sorted (not in original row order) so the exact per-home
                # sequence can't be inferred - only the distribution matters
                # for a median.
                'median': {k: sorted(v) for k, v in cell['median'].items()},
                'cat': {k: dict(v) for k, v in cell['cat'].items()},
            })
        year_modes_out[mode] = {
            'years': sorted([y for y in year_counts[mode] if isinstance(y, int)]),
            'cells': cells_out,
        }

    # Sorted (not in original row order) so the exact per-home sequence can't
    # be inferred - only the distribution matters for a median.
    median_export = {
        mode: {
            f'{city}|{year_key}': {key: sorted(values) for key, values in metrics.items()}
            for (city, year_key), metrics in by_city_year.items()
        }
        for mode, by_city_year in median_raw.items()
    }

    # Each home's own dict already has no fixed field order to preserve, and
    # sorting the dicts themselves would be meaningless - about all there is
    # to anonymize here is to not ship them in original row order.
    linked_home_export = {
        mode: {
            f'{city}|{year_key}': sorted(homes, key=lambda h: json.dumps(h, sort_keys=True))
            for (city, year_key), homes in by_city_year.items()
        }
        for mode, by_city_year in linked_homes_raw.items()
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
        'climate_zones': sorted(climate_zone_counts, key=lambda z: -climate_zone_counts[z]),
        'climate_zone_counts': climate_zone_counts,
        'step_achieved_counts': step_achieved_counts,
        'year_mode_default': 'built',
        'year_modes': year_modes_out,
        'median_export': median_export,
        'linked_home_export': linked_home_export,
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
