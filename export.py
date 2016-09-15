#!/usr/bin/env python
# encoding: utf-8

from __future__ import (absolute_import, division, print_function,
                        unicode_literals)

from decimal import Decimal
import re

from docx import Document


# Note: German technical terms (like "Gesamtergebnishaushalt") were not
# translated because they occur frequently in the original documents and
# translating them would have made that connection harder to understand.


# Structure of an "Ergebnishaushalt" table:
#
# The first row contains the header.
#
# Each following row has an optional running number, a "Kontogruppe"
# (except in the "Gesamtergebnishaushalt"), an optional sign (only if
# the row also has a running number), a title, and several value cells
# whose meaning is described in the corresponding header cell.
#
# Some records are split into multiple rows. In that case the first row
# has a running number, a sign, and value cells containing the sum of
# its sub-rows. Each sub-row has neither running number nor sign and
# its value cells contain individual values.
#
# Some records do not represent actual positions but instead summarize
# the previously listed records. These summary records can be recognized
# by having a "=" as their sign.



def split(s, maxsplit=None):
    '''
    Split a string at whitespace.

    Works like ``str.split`` with no explicit separator, i.e. splits
    the string ``s`` at any whitespace. In contrast to ``str.split``,
    however, you can set the maximum number of splits via ``maxsplit``
    while not having to pass an explicit separator.
    '''
    return re.split(r'\s+', s.strip(), maxsplit=maxsplit, flags=re.UNICODE)


def parse_amount(s):
    '''
    Parse a German amount string.
    '''
    parts = s.replace('.', '').split(',')
    if len(parts) == 1:
        parts.append('00')
    parts[1] = parts[1].ljust(2, '0')[:2]
    return Decimal('.'.join(parts))


def parse_int(s):
    '''
    Parse an int from a string, returns ``None`` for empty strings.
    '''
    if not s:
        return None
    return int(s)


KONTOGRUPPE_HEADER = 'Kto.\nGr.'


def extract_ergebnishaushalt(data):
    '''
    Extract an "Ergebnishaushalt" from tabular data.

    ``data`` is the data from a "Ergebnishaushalt" table as extracted by
    ``extract_table``. Both "Gesamtergebnishaushalt" and
    Teilergebnishaushalt" tables are supported.

    Returns a list of positions. Each position has a number, a sign, an
    optional "Kontogruppe", a title, a list of children and a list
    values. Each value has a type, a year, and an amount.

    The list of children may be empty. If it's not then the position's
    values are the sums of the corresponding values of its children. In
    that case, each child has a title and its own list of values.
    '''
    header = data[0]
    has_kontogruppe_column = header[1] == KONTOGRUPPE_HEADER
    years = []
    types = []
    if has_kontogruppe_column:
        first_value_column = 4
    else:
        first_value_column = 3
    for cell in header[first_value_column:]:
        type, year, _ = split(cell, 2)
        years.append(parse_int(year))
        types.append(type.lower())

    def parse_row(row):
        record = {'number': parse_int(row[0])}
        if has_kontogruppe_column:
            record['kontogruppe'] = parse_int(row[1])
            offset = 2
        else:
            record['kontogruppe'] = None
            offset = 1
        record['sign'] = row[0 + offset]
        record['title'] = row[1 + offset]
        record['values'] = values = []
        for i, cell in enumerate(row[2 + offset:]):
            values.append({'type': types[i], 'year': years[i],
                           'amount': parse_amount(cell)})
        return record

    positions = []
    position = None
    for row in data[2:]:  # The second row is part of the header
        record = parse_row(row)
        if record['number']:
            assert record['sign']
            record['children'] = []
            position = record
            positions.append(position)
        else:
            assert not record.pop('sign')
            assert not record.pop('kontogruppe')
            position['children'].append(record)
        assert position is not None
    return positions


def extract_table(table):
    data = []
    for row in table.rows:
        data.append([cell.text.strip() for cell in row.cells])
    return data

if __name__ == '__main__':
    import sys
    from pprint import pprint

    filename = sys.argv[1]
    doc = Document(filename)

    for table in doc.tables:
        data = extract_table(table)
        result = extract_gesamtergebnishaushalt(data)
        pprint(result)
        print('')
