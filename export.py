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


class Table(list):
    '''
    Base class for extracted tables.

    This class and its subclass encapsulate the parsing of the
    respective table layouts. The extracted data is represented as a
    list of positions, hence this class is a subclass of ``list``.

    Parsing is split over several methods to allow for customization by
    subclasses:

    - The entry point for parsing is ``_parse``, which is called by
      ``__init__`` with the table data. It first calls
      ``_parse_non_value_headers`` and ``_parse_value_headers`` before
      using ``_parse_row`` to parse the individual rows and group them
      into positions and their children.

    - ``_parse_non_value_headers`` must be implemented by subclasses. It
      receives the first row of the data and must set
      ``self._non_value_columns`` to a list of tuples. Each tuple
      corresponds to a non-value column and consists of a key under
      which the column's values are stored and a transform function for
      transforming the values (for example to parse numbers). The
      transform function can be ``None`` in which case the value is
      stored as text. The information from ``_non_value_columns`` is
      later used by ``_parse_row`` to identify and extract the data from
      these columns.

    - ``_parse_value_headers`` receives the remaining columns in the
      header and uses them to set ``self._types`` and ``self._years`` to
      lists containing the value column's types and years, respectively.
    '''

    def __init__(self, data):
        super(Table, self).__init__()
        self._parse(data)

    def _parse_non_value_headers(self, header):
        raise NotImplementedError('Must be implemented by subclass')

    def _parse_value_headers(self, headers):
        self._years = []
        self._types = []
        for cell in headers:
            type, year, _ = split(cell, 2)
            self._types.append(type.lower())
            self._years.append(parse_int(year))
        #print(self._types)
        #print(self._years)

    def _parse_row(self, row):
        values = []
        record = {'values': values}
        for i, (key, transform) in enumerate(self._non_value_columns):
            value = row[i]
            if transform:
                value = transform(value)
            record[key] = value
        for j, cell in enumerate(row[i + 1:]):
            values.append({'type': self._types[j], 'year': self._years[j],
                           'amount': parse_amount(cell)})
        return record

    def _parse(self, data):
        header = data[0]
        self._parse_non_value_headers(header)
        self._parse_value_headers(header[len(self._non_value_columns):])
        position = None
        for row in data[2:]:  # The second row is part of the header
            record = self._parse_row(row)
            if record['number']:
                assert record['sign']
                record['children'] = []
                position = record
                self.append(position)
            else:
                assert not record.pop('sign')
                assert not record.pop('kontogruppe', None)
                position['children'].append(record)
            assert position is not None


class ErgebnishaushaltTable(Table):

    _KONTOGRUPPE_HEADER = 'Kto.\nGr.'

    def _parse_non_value_headers(self, header):
        self._non_value_columns = [('number', parse_int)]
        if header[1] == self._KONTOGRUPPE_HEADER:
            self._non_value_columns.append(('kontogruppe', None))
        self._non_value_columns.extend([
            ('sign', None),
            ('title', None),
        ])

    def _parse_row(self, row):
        record = super(ErgebnishaushaltTable, self)._parse_row(row)
        # In the Gesamtergebnishaushalt there's no Kontogruppe, so we add the
        # field in case it's missing to get a consistent interface.
        record.setdefault('kontogruppe', None)
        return record


class FinanzhaushaltTable(Table):

    def _parse_non_value_headers(self, header):
        self._non_value_columns = [
            ('number', parse_int),
            ('sign', None),
            ('title', None)
        ]


def extract_table(table):
    data = []
    for row in table.rows:
        data.append([cell.text.strip() for cell in row.cells])
    return data


def convert_table(table):
    data = extract_table(table)
    if 'finanzhaushalt' in data[0][2]:
        return FinanzhaushaltTable(data)
    else:
        return ErgebnishaushaltTable(data)


if __name__ == '__main__':
    import sys
    from pprint import pprint

    filename = sys.argv[1]
    doc = Document(filename)

    for table in doc.tables:
        result = convert_table(table)
        pprint(result)
        print('')
