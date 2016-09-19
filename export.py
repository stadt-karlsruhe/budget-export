#!/usr/bin/env python
# encoding: utf-8

from __future__ import (absolute_import, division, print_function,
                        unicode_literals)

from decimal import Decimal
import re

from backports import csv
from docx import Document
from docx.table import Table as WordTable
from docx.text.paragraph import Paragraph


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


# Adapated from https://github.com/python-openxml/python-docx/issues/276
def iter_block_items(parent):
    '''
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    '''
    from docx.document import Document as _Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell

    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError('Unknown parent class {}'.format(
                         parent.__class__.__name__))

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield WordTable(child, parent)


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
      it on to ``_parse_meta_headers`` before calling
      ``_parse_value_headers``

    - ``_parse_meta_headers`` must be implemented by subclasses. It
      receives the first row of the data and must set
      ``self._meta_columns`` to a dict that maps column indices to
      2-tuples containing the key under which the column's values are
      stored and a transform function for transforming the values (for
      example to parse numbers). The transform function can be ``None``
      in which case the value is stored as text. The information from
      ``_meta_columns`` is later used by ``_parse_row`` to identify and
      extract the data from these columns.

    - ``_parse_value_headers`` receives the remaining columns in the
      header and uses them to set ``self._value_columns`` to a dict that
      maps column indices to 2-tuples containing the column's type and
      year.
    '''
    def __init__(self, data, teilhaushalt=None, produktbereich=None,
                 produktgruppe=None):
        super(Table, self).__init__()
        self._parse(data)
        self.teilhaushalt = teilhaushalt
        self.produktbereich = produktbereich
        self.produktgruppe = produktgruppe

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
                assert not record['sign']
                position['children'].append(record)
            assert position is not None

    def _csv_records(self):
        '''
        Yield records for dumping to CSV.

        This method allows subclasses whose rows do not follow the
        standard "Position + optional children" pattern to convert their
        rows into something suitable for CSV export.
        '''
        return iter(self)

    def dump_csv(self, writer, additional_columns=None, meta_columns=None,
                 include_summaries=False):
        if additional_columns is None:
            additional_columns = []
        if meta_columns is None:
            meta_columns = [c[0] for c in self._meta_columns.itervalues()]

        def dump_record(record, parent=None):
            fields = list(additional_columns)
            for key in meta_columns:
                value = record.get(key)
                if (not value) and (not include_summaries) and parent:
                    # Inherit from parent
                    value = parent.get(key)
                elif (key == 'title') and (not include_summaries) and parent:
                    value = '{}: {}'.format(parent['title'], value)
                fields.append(value)
            for value in record['values']:
                writer.writerow(fields + [value['year'], value['type'], value['amount']])

        for record in self._csv_records():
            if (record['sign'] == '=') and not include_summaries:
                continue
            if record['children']:
                if include_summaries:
                    dump_record(record)
                for child in record['children']:
                    dump_record(child, record)
            else:
                dump_record(record)


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

    def _csv_records(self):
        for project in self:
            for position in project['positions']:
                record = position.copy()
                record['project_id'] = project['id']
                record['project_title'] = project['title']
                yield record


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


class _HeadingState(object):
    '''
    A state machine for tracking the information from the headings.

    The headings contain information on which Teilhaushalt,
    Produktbereich, and Produktgruppe a table belongs to. However,
    the structure of the headings does not reflect that directly.
    Instead, their meaning has to be inferred from their order and
    their content. This class contains the necessary logic for this.
    '''
    def __init__(self):
        self.reset()
        self.teilhaushalte = {}

    def reset(self):
        self.teilhaushalt = None
        self.produktbereich = None
        self.produktgruppe = None

    def register_heading(self, text):
        '''
        Register a heading from the document.

        ``text`` is the text of the heading.
        '''
        text = text.strip()
        if not text:
            return
        parts = split(text, 1)
        if len(parts) != 2:
            return
        id, title = parts
        if id.startswith('THH'):
            id = id[3:]
            self.teilhaushalt = self.teilhaushalte.setdefault(
                    id, {'id': id, 'title': title, 'produktbereiche': {}})
            self.produktbereich = None
            self.produktgruppe = None
        elif (
            self.teilhaushalt is not None
            and self.produktbereich is None
            and len(id) == 2
            and id.isdigit()
        ):
            self.produktbereich = self.teilhaushalt['produktbereiche'].setdefault(
                    id, {'id': id, 'title': title, 'produktgruppen': {}})
            self.produktgruppe = None
        elif (
            self.produktbereich is not None
            and self.produktgruppe is None
            and len(id) == 4
            and id.isdigit()
        ):
            self.produktgruppe = self.produktbereich['produktgruppen'].setdefault(
                    id, {'id': id, 'title': title})


if __name__ == '__main__':
    import io
    import sys
    from pprint import pprint

    filenames = sys.argv[1:]

    headings = _HeadingState()

    tables = []
    for filename in filenames:
        print('Loading {}'.format(filename))
        doc = Document(filename)
        headings.reset()
        for block in iter_block_items(doc):
            if isinstance(block, WordTable):
                sys.stdout.write('.')
                sys.stdout.flush()
                data = extract_data(block)
                try:
                    table = convert_data(data)
                except ValueError:
                    # Assume it's a sub-table of an Investitionsübersicht
                    tables[-1].append_data(data)
                else:
                    if headings.teilhaushalt:
                        table.teilhaushalt = headings.teilhaushalt['id']
                    if headings.produktbereich:
                        table.produktbereich = headings.produktbereich['id']
                    if headings.produktgruppe:
                        table.produktgruppe = headings.produktgruppe['id']
                    tables.append(table)
            else:
                headings.register_heading(block.text)
        print('')


    print('')
    print('-' * 70)
    print('')
    pprint(headings.teilhaushalte)
    print('')
    print('-' * 70)
    print('')

    for table in tables:
        pprint(table)
        print('')
        print('-' * 70)
        print('')

    # Fix https://github.com/ryanhiebert/backports.csv/issues/14
    import numbers
    csv.number_types = (numbers.Number,)

    csv_options = {'delimiter': ';', 'quoting': csv.QUOTE_NONNUMERIC}

    def dump_csv(filename, table_filter, header, meta_columns, additional_fields=None):
        '''
        Dump tables to a CSV file.

        ``filename`` is the name of the file.

        ``table_filter`` is a callback that gets a ``Table`` instance
        and returns ``True`` if the table should be dumped.

        ``header`` is a list of header labels for the first row of the
        CSV file.

        ``meta_columns`` is a list of the keys of the meta column which
        should be exported from the table. Note that value columns are
        always exported.

        ``additional_fields`` is an optional callback that gets a a
        ``Table`` instance and returns a list of additional fields.
        These fields are prefixed to the fields of each row in the
        table.
        '''
        with io.open(filename, 'w') as f:
            writer = csv.writer(f, **csv_options)
            writer.writerow(header)
            for table in tables:
                if table_filter(table):
                    if additional_fields:
                        add_cols = additional_fields(table)
                    else:
                        add_cols = None
                    table.dump_csv(writer, meta_columns=meta_columns,
                                   additional_columns=add_cols)

    dump_csv('gesamtergebnishaushalt.csv',
             lambda t: isinstance(t, ErgebnishaushaltTable) and not t.teilhaushalt,
             ['TITEL', 'JAHR', 'TYP', 'BETRAG'],
             ['title'])

    dump_csv('teilergebnishaushalte.csv',
             lambda t: isinstance(t, ErgebnishaushaltTable) and t.teilhaushalt,
             ['TEILHAUSHALT', 'PRODUKTBEREICH', 'PRODUKTGRUPPE', 'KONTOGRUPPE',
              'TITEL', 'JAHR', 'TYP', 'BETRAG'],
             ['kontogruppe', 'title'],
             lambda t: [t.teilhaushalt, t.produktbereich, t.produktgruppe])

    dump_csv('gesamtfinanzshaushalt.csv',
             lambda t: isinstance(t, FinanzhaushaltTable) and not t.teilhaushalt,
             ['TITEL', 'JAHR', 'TYP', 'BETRAG'],
             ['title'])

    dump_csv('teilfinanzhaushalte.csv',
             lambda t: isinstance(t, FinanzhaushaltTable) and t.teilhaushalt,
             ['TEILHAUSHALT', 'TITEL', 'JAHR', 'TYP', 'BETRAG'],
             ['title'],
             lambda t: [t.teilhaushalt])

    dump_csv('investitionsuebersichten.csv',
             lambda t: isinstance(t, InvestitionsuebersichtTable),
             ['TEILHAUSHALT', 'PROJEKTNUMMER', 'PROJEKT', 'TITEL', 'JAHR',
              'TYP', 'BETRAG'],
             ['project_id', 'project_title', 'title'],
             lambda t: [t.teilhaushalt])
