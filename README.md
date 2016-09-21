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


## Installation

To install the required Python packages (we recommend using a
[virtualenv][virtualenv]):

    pip install -r requirements.txt

[virtualenv]: https://virtualenv.pypa.io


## Usage

The script `export.py` takes one or more filenames of Word documents as
parameters and writes the exported data to CSV files in the current directory:

    python export.py word_document1.docx word_document2.docx

If the export from KM-Doppik is split into multiple Word documents then all of
these should be passed to `export.py` in a single run.

