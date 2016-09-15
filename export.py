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
      ``__init__`` with the table data. It calls ``_parse_headers``
      before using ``_parse_row`` to parse the individual rows and group
      them into positions and their children.

    - ``_parse_headers`` receives the first row of the table and passes
      it on to ``_parse_non_value_headers`` before calling ``_parse_value_headers``

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

    def _parse_meta_headers(self, header):
        raise NotImplementedError('Must be implemented in subclass')

    def _parse_value_headers(self, header):
        self._value_columns = {}
        num_meta = len(self._meta_columns)
        for i, cell in enumerate(header[num_meta:], num_meta):
            parts = cell.split()
            if (len(parts) == 3) and (parts[2] == 'EUR'):
                try:
                    year = parse_int(parts[1])
                except ValueError:
                    # Not a year
                    continue
                self._value_columns[i] = (parts[0], year)

    def _parse_headers(self, header):
        self._parse_meta_headers(header)
        self._parse_value_headers(header)

    def _parse_row(self, row):
        values = []
        record = {'values': values}
        for i, (key, transform) in self._meta_columns.iteritems():
            value = row[i]
            if transform:
                value = transform(value)
            record[key] = value
        for i, (type, year) in self._value_columns.iteritems():
            values.append({'type': type, 'year': year,
                           'amount': parse_amount(row[i])})
        return record

    def _parse(self, data):
        self._parse_headers(data[0])
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

    def _parse_meta_headers(self, header):
        self._meta_columns = {
            0: ('number', parse_int),
        }
        if header[1] == self._KONTOGRUPPE_HEADER:
            self._meta_columns[1] = ('kontogruppe', None)
            offset = 1
        else:
            offset = 0
        self._meta_columns[1 + offset] = ('sign', None)
        self._meta_columns[2 + offset] = ('title', None)

    def _parse_row(self, row):
        record = super(ErgebnishaushaltTable, self)._parse_row(row)
        # In the Gesamtergebnishaushalt there's no Kontogruppe, so we add the
        # field in case it's missing to get a consistent interface.
        record.setdefault('kontogruppe', None)
        return record


class FinanzhaushaltTable(Table):

    def _parse_meta_headers(self, header):
        self._meta_columns = {
            0: ('number', parse_int),
            1: ('sign', None),
            2: ('title', None),
        }


class InvestitionsuebersichtTable(Table):

    def _parse_meta_headers(self, header):
        self._meta_columns = {
            0: ('number', parse_int),
            1: ('sign', None),
            2: ('title', None),
        }

    def _parse(self, data):
        self._parse_headers(data[0])
        self._parse_body(data[2:])  # Second row is part of the header

    def _parse_body(self, rows):
        project = None
        position = None
        for row in rows:
            if len(set(row)) == 1:
                # Row consisting of a single merged cell containing the
                # project ID and name.
                parts = row[0].split(':', 1)
                project = {
                    'id': parts[0].strip(),
                    'title': parts[1].strip(),
                    'positions': [],
                }
                self.append(project)
            else:
                # Standard row
                record = self._parse_row(row)
                if record['number']:
                    assert record['sign']
                    record['children'] = []
                    position = record
                    project['positions'].append(position)
                else:
                    assert not record.pop('sign')
                    assert not record.pop('kontogruppe', None)
                    position['children'].append(record)

    def append_data(self, data):
        '''
        Add more data for this table.

        If a Investitionsübersicht contains multiple projects then each
        project is exported into a separate Word table. Only the first
        of these has a header and should be parsed by creating a new
        ``InvestitionsuebersichtTable`` instance. The other, header-less
        tables then can be added to that table via ``append_data``.
        '''
        self._parse_body(data)


def extract_data(table):
    data = []
    for row in table.rows:
        data.append([cell.text.strip() for cell in row.cells])
    return data


def convert_data(data):
    if 'finanzhaushalt' in data[0][2].lower():
        return FinanzhaushaltTable(data)
    elif 'investitionsübersicht' in data[0][2].lower():
        return InvestitionsuebersichtTable(data)
    elif ('ergebnishaushalt' in data[0][2].lower()
            or 'ergebnishaushalt' in data[0][3].lower()):
        return ErgebnishaushaltTable(data)
    else:
        raise ValueError('Unknown table type.')


if __name__ == '__main__':
    import sys
    from pprint import pprint

    filenames = sys.argv[1:]

    tables = []
    for filename in filenames:
        print('Loading {}'.format(filename))
        doc = Document(filename)
        for table in doc.tables:
            sys.stdout.write('.')
            sys.stdout.flush()
            data = extract_data(table)
            try:
                tables.append(convert_data(data))
            except ValueError:
                # Assume it's a sub-table of an Investitionsübersicht
                tables[-1].append_data(data)
        print('')

    for table in tables:
        pprint(table)
        print('')
        print('-' * 70)
        print('')

