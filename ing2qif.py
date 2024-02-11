#!/usr/bin/env python3
# (C) 2014, Marijn Vriens <marijn@metronomo.cl>
# GNU General Public License, version 3 or any later version

# Documents
# https://github.com/Gnucash/gnucash/blob/master/src/import-export/qif-imp/file-format.txt
# https://en.wikipedia.org/wiki/Quicken_Interchange_Format

import csv
import itertools
import argparse

class Entry:
    """
    Represents one entry.
    """
    def __init__(self, data):
        self._data = data
        self._clean_up()

    def _clean_up(self):
        self._data['amount'] = self._data['Bedrag (EUR)'].replace(',', '.')

    def keys(self):
        return self._data.keys()

    def __getattr__(self, item):
        return self._data.get(item)

    def __getitem__(self, item):
        return self._data.get(item)


class CsvEntries:
    def __init__(self, file_descriptor):
        self._entries = csv.DictReader(file_descriptor, delimiter=';')

    def __iter__(self):
        return map(Entry, self._entries)


class QifEntries:
    def __init__(self):
        self._entries = []

    def add_entry(self, entry):
        """
        Add an entry to the list of entries in the statment.
        :param entry: A dictionary where each key is one of the keys of the statement.
        :return: Nothing.
        """
        self._entries.append(QifEntry(entry))

    def serialize(self):
        """
        Turn all the entries into a string
        :return: a string with all the entries.
        """
        data = ["!Type:Bank"]
        for e in self._entries:
            data.append(e.serialize())
        return "\n".join(data)


class QifEntry:
    def __init__(self, entry):
        self._entry = entry
        self._data = []
        self._processing()

    def _processing(self):
        self._data.append("D{}".format(self._entry.Datum))
        self._data.append("T{}".format(self._amount_format()))
        entry_type = self._entry_type()
        if entry_type:
            self._data.append('N{}'.format(entry_type))
        self._data.append("M{}".format(self._memo()))
        self._data.append("^")

    def serialize(self):
        """
        Turn the QifEntry into a String.
        :return: a string
        """
        return "\n".join(self._data)

    def _memo_geldautomaat(self, mededelingen, omschrijving):
        if omschrijving.startswith('ING>') or \
                omschrijving.startswith('ING BANK>') or \
                omschrijving.startswith('OPL. CHIPKNIP'):
            memo = omschrijving
        else:
            memo = mededelingen[:32]
        return memo

    def _memo_incasso(self, mededelingen, omschrijving):
        if omschrijving.startswith('SEPA Incasso') or mededelingen.startswith('SEPA Incasso'):
            try:
                s = mededelingen.index('Naam: ')+6
            except:
                raise Exception(mededelingen, omschrijving)
            e = mededelingen.index('Kenmerk: ')
            return  mededelingen[s:e]

    def _memo_internetbankieren(self, mededelingen, omschrijving):
        try:
            s = mededelingen.index('Naam: ')+6
            if "Omschrijving:" in mededelingen:
                e = mededelingen.index('Omschrijving: ')
            else:
                e = mededelingen.index('IBAN: ')
            return  mededelingen[s:e]
        except ValueError:
            return None

    def _memo_diversen(self, mededelingen, omschrijving):
        return mededelingen[:64]

    def _memo_verzamelbetaling(self, mededelingen, omschrijving):
        if 'Naam: ' in mededelingen:
            s = mededelingen.index('Naam: ')+6
            e = mededelingen.index('Kenmerk: ')
            return  mededelingen[s:e]

    def _memo(self):
        """
        Decide what the memo field should be. Try to keep it as sane as possible. If unknown type, include all data.
        :return: the memo field.
        """
        mutatie_soort = self._entry['MutatieSoort']
        mededelingen = self._entry['Mededelingen']
        omschrijving = self._entry['Naam / Omschrijving']

        memo = None
        try:
            memo_method = { # Depending on the mutatie_soort, switch memo generation method.
                'Diversen': self._memo_diversen,
                'Betaalautomaat': self._memo_geldautomaat,
                'Geldautomaat': self._memo_geldautomaat,
                'Incasso': self._memo_incasso,
                'Internetbankieren': self._memo_internetbankieren,
                'Overschrijving': self._memo_internetbankieren,
                'Verzamelbetaling': self._memo_verzamelbetaling,
            }[mutatie_soort]
            memo = memo_method(mededelingen, omschrijving)
        except KeyError:
            pass
        finally:
            if memo is None:
                # The default memo value. All the text.
                memo = "%s %s" % (self._entry['Mededelingen'], self._entry['Naam / Omschrijving'])
        if self._entry_type():
            return "%s %s" % (self._entry_type(), memo)
        return memo.strip()


    def _amount_format(self):
        if self._entry['Af Bij'] == 'Bij':
            return "+" + self._entry['amount']
        else:
            return "-" + self._entry['amount']

    def _entry_type(self):
        """
        Detect the type of entry.
        :return:
        """
        try:
            return {
                'Geldautomaat': "ATM",
                'Internetbankieren': "Transfer",
                'Incasso': 'Transfer',
                'Verzamelbetaling': 'Transfer',
                'Betaalautomaat': "ATM",
                'Storting': 'Deposit',
            }[self._entry['MutatieSoort']]
        except KeyError:
            return None


def main(file_descriptor, start, number):
    qif = QifEntries()
    for c, entry in enumerate(CsvEntries(file_descriptor), 1):
        if c >= start:
            qif.add_entry(entry)
            if number and c > start + number - 1:
                break
    print(qif.serialize())

def parse_cmdline():
    parser = argparse.ArgumentParser(description="Convert CSV banking statements to QIF format for GnuCash.")
    parser.add_argument("csvfile", metavar="CSV_FILE", help="The CSV file with banking statements.")
    parser.add_argument("--start", type=int, metavar="NUMBER", default=1, help="The statement to start conversion at.")
    parser.add_argument("--number", type=int, metavar="NUMBER", help="The number of statements to convert.")
    return parser.parse_args()

if __name__ == '__main__':
    args = parse_cmdline()
    with open(args.csvfile, 'r', encoding='utf-8') as fd:
        main(fd, args.start, args.number)
