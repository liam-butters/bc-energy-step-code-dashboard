"""
Maps a CLIENTPCODE's Forward Sortation Area (FSA - the first 3 characters of
a Canadian postal code, e.g. "V8N") to the municipality it belongs to, so the
dashboard can bucket a row by its postal code instead of its self-reported
(and often typo'd or inconsistent) CLIENTCITY text.

Built empirically: for every FSA appearing in the workbook, this is the
municipality name most self-reported CLIENTCITY entries with that FSA agree
on (after the same typo-fixing extract_data.normalize_city applies), cross-
checked against Canada Post's own FSA area descriptions (Wikipedia's "List of
postal codes of Canada: V"). The two agree almost everywhere - an FSA's
"catchment" rarely splits across municipal lines outside the densely-packed
Capital Regional District, where several FSAs genuinely straddle more than
one real municipality (per Canada Post's own description) and no single
"correct" answer exists; those are left on the self-reported majority too.

One explicit override: V8N/V8P/V8X/V8Y/V8Z/V9E are Saanich's official FSAs
(per Canada Post's own area description, listing "Saanich" in the sub-area
for each) even though a few of them show "Victoria" as the self-reported
majority - e.g. UVic (in V8P) and other Saanich neighbourhoods are commonly
self-reported as "Victoria" colloquially.

Necessarily approximate for the same reasons regions.py is: a handful of
very-low-volume FSAs (fewer than 5 rows) have a "majority" that's really just
a data-entry error in one row's postal code - not worth hand-curating given
how little they affect the totals.

Regenerate by re-running the FSA/city cross-tab against a fresh workbook
export if the data changes meaningfully (see dashboard dev notes).
"""
import re

FSA_RE = re.compile(r'^([A-Za-z]\d[A-Za-z])')

FSA_TO_CITY = {
    'V0A': 'Invermere',  # 10/26
    'V0B': 'Creston',  # 59/299
    'V0C': 'Rose Prairie',  # 4/7
    'V0E': 'Revelstoke',  # 77/244
    'V0G': 'Rossland',  # 84/194
    'V0H': 'Summerland',  # 99/334
    'V0J': 'Smithers',  # 32/64
    'V0K': '100 Mile House',  # 36/104
    'V0L': 'Hanceville',  # 1/1
    'V0M': 'Agassiz',  # 36/41
    'V0N': 'Gibsons',  # 115/470
    'V0P': 'Kelowna',  # 1/3 (very low volume; likely a data-entry error, see module note)
    'V0R': 'Tofino',  # 52/336
    'V0S': 'Sooke',  # 3/3
    'V0T': 'Skidegate',  # 1/1
    'V0V': 'Delta',  # 2/4
    'V0X': 'Princeton',  # 51/101
    'V1A': 'Kimberley',  # 194/207
    'V1B': 'Vernon',  # 115/167
    'V1C': 'Cranbrook',  # 235/239
    'V1E': 'Salmon Arm',  # 72/77
    'V1H': 'Vernon',  # 168/171
    'V1J': 'Fort St. John',  # 14/14
    'V1K': 'Merritt',  # 1/1
    'V1L': 'Nelson',  # 168/187
    'V1M': 'Langley',  # 96/106
    'V1N': 'Castlegar',  # 96/107
    'V1P': 'Kelowna',  # 211/216
    'V1R': 'Trail',  # 10/12
    'V1S': 'Kamloops',  # 23/27
    'V1T': 'Vernon',  # 47/48
    'V1V': 'Kelowna',  # 530/532
    'V1W': 'Kelowna',  # 346/348
    'V1X': 'Kelowna',  # 96/97
    'V1Y': 'Kelowna',  # 281/284
    'V1Z': 'West Kelowna',  # 46/75
    'V2A': 'Penticton',  # 395/405
    'V2B': 'Kamloops',  # 35/39
    'V2C': 'Kamloops',  # 99/100
    'V2E': 'Kamloops',  # 70/71
    'V2G': 'Williams Lake',  # 8/8
    'V2H': 'Kamloops',  # 11/11
    'V2J': 'Quesnel',  # 39/44
    'V2K': 'Prince George',  # 145/151
    'V2L': 'Prince George',  # 19/21
    'V2M': 'Prince George',  # 20/21
    'V2N': 'Prince George',  # 215/221
    'V2P': 'Chilliwack',  # 45/46
    'V2R': 'Chilliwack',  # 138/143
    'V2S': 'Abbotsford',  # 92/92
    'V2T': 'Abbotsford',  # 128/131
    'V2V': 'Mission',  # 18/18
    'V2W': 'Maple Ridge',  # 68/68
    'V2X': 'Maple Ridge',  # 12/14
    'V2Y': 'Langley',  # 1463/1472
    'V2Z': 'Langley',  # 152/170
    'V3A': 'Langley',  # 150/155
    'V3B': 'Port Coquitlam',  # 64/107
    'V3C': 'Port Coquitlam',  # 31/45
    'V3E': 'Coquitlam',  # 166/172
    'V3G': 'Abbotsford',  # 69/69
    'V3H': 'Port Moody',  # 55/76
    'V3J': 'Coquitlam',  # 49/52
    'V3K': 'Coquitlam',  # 64/69
    'V3L': 'New Westminster',  # 64/64
    'V3M': 'New Westminster',  # 73/78
    'V3N': 'Burnaby',  # 108/108
    'V3R': 'Surrey',  # 89/89
    'V3S': 'Surrey',  # 558/596
    'V3T': 'Surrey',  # 58/65
    'V3V': 'Surrey',  # 168/173
    'V3W': 'Surrey',  # 108/192
    'V3X': 'Surrey',  # 140/140
    'V3Y': 'Pitt Meadows',  # 14/15
    'V3Z': 'Surrey',  # 218/224
    'V4A': 'Surrey',  # 154/156
    'V4B': 'White Rock',  # 6/6
    'V4C': 'Delta',  # 24/27
    'V4E': 'Delta',  # 5/5
    'V4K': 'Delta',  # 28/29
    'V4L': 'Delta',  # 38/38
    'V4M': 'Tsawwassen',  # 41/61
    'V4N': 'Surrey',  # 365/369
    'V4P': 'Surrey',  # 24/27
    'V4R': 'Maple Ridge',  # 11/11
    'V4S': 'Mission',  # 10/10
    'V4T': 'West Kelowna',  # 134/155
    'V4V': 'Lake Country',  # 255/443
    'V4W': 'Langley',  # 52/65
    'V4X': 'Abbotsford',  # 32/32
    'V4Y': 'Armstrong',  # 2/5
    'V4Z': 'Chilliwack',  # 19/19
    'V5A': 'Burnaby',  # 84/84
    'V5B': 'Burnaby',  # 122/123
    'V5C': 'Burnaby',  # 96/97
    'V5E': 'Burnaby',  # 86/89
    'V5G': 'Burnaby',  # 117/117
    'V5H': 'Burnaby',  # 46/46
    'V5J': 'Burnaby',  # 84/85
    'V5K': 'Vancouver',  # 359/361
    'V5L': 'Vancouver',  # 130/131
    'V5M': 'Vancouver',  # 283/284
    'V5N': 'Vancouver',  # 337/337
    'V5P': 'Vancouver',  # 332/332
    'V5R': 'Vancouver',  # 398/403
    'V5S': 'Vancouver',  # 256/257
    'V5T': 'Vancouver',  # 99/101
    'V5V': 'Vancouver',  # 264/264
    'V5W': 'Vancouver',  # 256/257
    'V5X': 'Vancouver',  # 319/321
    'V5Y': 'Vancouver',  # 145/146
    'V5Z': 'Vancouver',  # 159/159
    'V6A': 'Vancouver',  # 6/6
    'V6C': 'Squamish',  # 1/1
    'V6G': 'Vancouver',  # 4/4
    'V6H': 'Vancouver',  # 33/33
    'V6J': 'Vancouver',  # 42/42
    'V6K': 'Vancouver',  # 101/101
    'V6L': 'Vancouver',  # 116/116
    'V6M': 'Vancouver',  # 127/127
    'V6N': 'Vancouver',  # 145/145
    'V6P': 'Vancouver',  # 228/228
    'V6R': 'Vancouver',  # 199/199
    'V6S': 'Vancouver',  # 101/102
    'V6T': 'Vancouver',  # 7/8
    'V6V': 'Richmond',  # 12/12
    'V6W': 'Richmond',  # 6/6
    'V6X': 'Richmond',  # 47/47
    'V6Y': 'Richmond',  # 70/72
    'V7A': 'Richmond',  # 179/179
    'V7B': 'Richmond',  # 16/16
    'V7C': 'Richmond',  # 201/203
    'V7E': 'Richmond',  # 222/222
    'V7G': 'North Vancouver',  # 9/9
    'V7H': 'North Vancouver',  # 4/4
    'V7J': 'North Vancouver',  # 34/34
    'V7K': 'North Vancouver',  # 11/13
    'V7L': 'North Vancouver',  # 117/125
    'V7M': 'North Vancouver',  # 33/35
    'V7N': 'North Vancouver',  # 23/25
    'V7P': 'North Vancouver',  # 51/52
    'V7R': 'North Vancouver',  # 53/57
    'V7S': 'West Vancouver',  # 58/58
    'V7T': 'West Vancouver',  # 24/24
    'V7V': 'West Vancouver',  # 58/59
    'V7W': 'West Vancouver',  # 20/21
    'V7Z': 'Sechelt',  # 60/68
    'V8A': 'Powell River',  # 61/62
    'V8B': 'Squamish',  # 214/214
    'V8C': 'Kitimat',  # 1/1
    'V8E': 'Whistler',  # 53/53
    'V8G': 'Terrace',  # 3/3
    'V8H': 'Mill Bay',  # 6/14
    'V8K': 'Salt Spring Island',  # 9/10
    'V8L': 'North Saanich',  # 121/199 (straddles Sidney/North Saanich - see module note)
    'V8M': 'Central Saanich',  # 62/90
    'V8N': 'Saanich',
    'V8P': 'Saanich',  # official Saanich FSA (self-reported majority was Victoria 33/59 - UVic etc.)
    'V8R': 'Victoria',  # 35/50 (straddles Victoria/Oak Bay - see module note)
    'V8S': 'Victoria',  # 61/90 (straddles Victoria/Oak Bay - see module note)
    'V8T': 'Victoria',  # 18/18
    'V8V': 'Victoria',  # 21/22
    'V8X': 'Saanich',  # official Saanich FSA (self-reported majority was Victoria 53/85)
    'V8Y': 'Saanich',  # official Saanich FSA (self-reported majority was Victoria 21/37)
    'V8Z': 'Saanich',
    'V9A': 'Victoria',  # 47/61 (straddles Victoria/Esquimalt - see module note)
    'V9B': 'Langford',  # 449/713 (straddles Highlands/Langford/View Royal - see module note)
    'V9C': 'Colwood',  # 481/641 (straddles Colwood/Langford/Metchosin - see module note)
    'V9E': 'Saanich',  # official Saanich FSA (self-reported majority was Victoria 5/9)
    'V9G': 'Ladysmith',  # 85/87
    'V9H': 'Campbell River',  # 64/64
    'V9J': 'Courtenay',  # 27/47
    'V9K': 'Qualicum Beach',  # 40/57
    'V9L': 'Duncan',  # 104/210
    'V9M': 'Comox',  # 149/162
    'V9N': 'Courtenay',  # 424/441
    'V9P': 'Parksville',  # 115/140
    'V9R': 'Nanaimo',  # 259/264
    'V9S': 'Nanaimo',  # 43/47
    'V9T': 'Nanaimo',  # 268/276
    'V9V': 'Nanaimo',  # 18/22
    'V9W': 'Campbell River',  # 170/170
    'V9X': 'Nanaimo',  # 46/46
    'V9Y': 'Port Alberni',  # 90/92
    'V9Z': 'Sooke',  # 199/205
}


def city_from_fsa(raw):
    """Returns the mapped municipality for raw's FSA, or None if raw is blank,
    not a recognizable postal code, or its FSA isn't in FSA_TO_CITY (e.g. one
    that's never appeared in the workbook before)."""
    if not raw:
        return None
    m = FSA_RE.match(raw.strip().replace(' ', ''))
    if not m:
        return None
    return FSA_TO_CITY.get(m.group(1).upper())
