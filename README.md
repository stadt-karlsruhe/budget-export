# Export of Budget Data from Word to CSV

This script exports budget data from Word documents to machine-readable CSV
files. It is intended for Word documents produced by the
["Ein-Knopf-Lösung"][onebutton] ("one-button-solution") export functionality
of the [KM-Doppik][kmdoppik] software by the [Datenzentrale][datenzentrale].

The supported formats are

* Gesamtergebnishaushalt (total profit and loss budget)
* Gesamtfinanzhaushalt (total cash-flow budget)
* Teilergebnishaushalte (partial profit and loss budgets)
* Teilfinanzhaushalte (partial cash-flow budgets)
* Investitionsübersichten (summaries of investments)

[onebutton]: http://www.datenzentrale.de/,Lde/Start/Die+Loesungen/KM-Doppik_FormatiertesReporting.html
[kmdoppik]: http://www.datenzentrale.de/,Lde/Start/Die+Loesungen/KM-Doppik.html
[datenzentrale]: http://www.datenzentrale.de

The script was developed for the budget of the City of Karlsruhe, using it with
other cities' data may require additional work. Feel free to contact us at
transparenz@karlsruhe.de if you run into problems.


## Installation

First clone this repository:

    git clone https://github.com/stadt-karlsruhe/budget-export.git

Then install the required Python packages (we recommend using a
[virtualenv][virtualenv]):

    pip install -r requirements.txt

[virtualenv]: https://virtualenv.pypa.io


## Usage

The script `budget_export.py` takes one or more filenames of Microsoft Word
`.docx` documents as parameters and writes the exported data to CSV files in
the current directory:

    python budget_export.py word_document1.docx word_document2.docx

If the export from KM-Doppik is split into multiple Word documents then all of
these should be passed to `budget_export.py` in a single run.


## Output format

The data is exported as UTF-8 encoded CSV files. Columns are separated by
`,` and strings are quoted using `"`. The first row in each file is a
header containing the column labels.

The data usually contains multiple values for a single position, for example
the planned and actual values for several years. To ease the combination of
data from several datasets these multiple values are mapped to multiple rows.

Summary rows that only aggregate the data of other rows are not exported.

Monetary amounts are in EUR (€). The decimal mark is `.`, no thousands
separator is used. Positive and negative values represent earnings and
expenses, correspondingly.

[Sample output from the City of Karlsruhe](https://transparenz.karlsruhe.de/dataset/haushaltsplan-daten-2017-2018)


## License

Copyright (c) 2017, Stadt Karlsruhe (www.karlsruhe.de)

Distributed under the MIT license, see the file `LICENSE` for details.


## Changes

# 0.1.0

* First public release

