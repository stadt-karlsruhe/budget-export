#!/usr/bin/env python
# encoding: utf-8

from __future__ import (absolute_import, division, print_function,
                        unicode_literals)

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


def extract_gesamtergebnishaushalt(data):
    '''
    Extract a "Gesamtergebnishaushalt" from tabular data.

    ``data`` is the data from a "Gesamtergebnishaushalt" table as
    extracted by ``extract_table``.

    Returns a list of positions. Each position has a number, a sign, a
    title, a list of children and a list values. Each value has a type,
    a year, and an actual value.

    The list of children may be empty. If it's not then the position's
    values are the sums of the corresponding values of its children. In
    that case, each child has a title and its own list of values.
    '''
    header = data[0]
    years = []
    types = []
    for cell in header[3:]:
        type, year, _ = split(cell, 2)
        years.append(year)
        types.append(type.lower())
    positions = []
    position = None
    for row in data[2:]:  # The second row is part of the header
        number = row[0]
        if number:
            position = {'number': number, 'sign': row[1], 'children': []}
            positions.append(position)
            record = position
        else:
            record = {}
            position['children'].append(record)
        assert position is not None
        record['title'] = row[2]
        record['values'] = values = []
        for i, cell in enumerate(row[3:]):
            values.append({'type': types[i], 'year': years[i], 'value': cell})
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
