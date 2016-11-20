import pprint
import sys
import zipfile

from enum import Enum, unique

import xlrd
from xlrd import X12Book
from xlrd import XLRDError
from xlrd import book
from xlrd import timemachine

@unique
class XmlOperation(Enum):
    openAndClose = 1
    open = 2
    close = 3

class ExcelReader:
    password = 'Florian112'
    sheetName = 'Stammliste neu'
    mapCellType = {0: lambda x: str(x),
                   1: lambda x: str(x),
                   2: lambda x: str(int(x)),
                   3: lambda x: str(x)}


    def __init__(self, filename):
        self.filename = filename

        print('operate on: ' + filename)

    def open_workbook(filename=None,
                      logfile=sys.stdout,
                      verbosity=0,
                      use_mmap=xlrd.USE_MMAP,
                      file_contents=None,
                      encoding_override=None,
                      formatting_info=False,
                      on_demand=False,
                      ragged_rows=False,
                      ):
        peeksz = 4
        if file_contents:
            peek = file_contents[:peeksz]
        else:
            with open(filename, "rb") as f:
                peek = f.read(peeksz)

        isZipFile = zipfile.is_zipfile(filename)
        zip = zipfile.ZipFile(filename)

        #zip.open(self.password)

        if peek == b"PK\x03\x04":  # a ZIP file
            if file_contents:
                zf = zipfile.ZipFile(timemachine.BYTES_IO(file_contents))
            else:
                zf = zipfile.ZipFile(filename)

            # Workaround for some third party files that use forward slashes and
            # lower case names. We map the expected name in lowercase to the
            # actual filename in the zip container.
            component_names = dict([(X12Book.convert_filename(name), name)
                                    for name in zf.namelist()])

            if verbosity:
                logfile.write('ZIP component_names:\n')
                pprint.pprint(component_names, logfile)
            if 'xl/workbook.xml' in component_names:
                from . import xlsx
                bk = xlsx.open_workbook_2007_xml(
                    zf,
                    component_names,
                    logfile=logfile,
                    verbosity=verbosity,
                    use_mmap=use_mmap,
                    formatting_info=formatting_info,
                    on_demand=on_demand,
                    ragged_rows=ragged_rows,
                )
                return bk
            if 'xl/workbook.bin' in component_names:
                raise XLRDError('Excel 2007 xlsb file; not supported')
            if 'content.xml' in component_names:
                raise XLRDError('Openoffice.org ODS file; not supported')
            raise XLRDError('ZIP file contents not a known type of workbook')

        bk = book.open_workbook_xls(
            filename=filename,
            logfile=logfile,
            verbosity=verbosity,
            use_mmap=use_mmap,
            file_contents=file_contents,
            encoding_override=encoding_override,
            formatting_info=formatting_info,
            on_demand=on_demand,
            ragged_rows=ragged_rows,
        )
        return bk

    def __getRelevantRows(self, sheet):
        rows = sheet.get_rows()

        relevantRows = []

        for row in rows:
            cellValue = row[0].value

            if cellValue == '' or cellValue is None:
                continue

            try:
                intValue = int(cellValue)
            except:
                continue
                pass

            if intValue is not None and intValue > 0:
                relevantRows.append(row)

        return relevantRows

    def __convertRowsToXml(self, rows):
        xml = "<?xml version='1.0' encoding='UTF-8'?>\n" \
              "<items>"

        for row in rows:
            rowAsXml = self.__convertRowToXml(row)
            xml += "\n" + rowAsXml

        xml += "\n</items>"

        return xml

    def __convertRowToXml(self, row):
        offsetLp = 36
        offsetAusbildung = offsetLp + 12 + 7
        offsetVerfuegbarkeit = offsetAusbildung + 8 + 1

        ids = {0: "Nummer", 2: "aktivUeber18", 3: "aktivUnter18", 4: "maennlich", 5: "weiblich",
               7: "vereinAktiv", 9: "vereinPassiv", 11: "vereinFoerdernd",
               14: "rang", 15: "gruppe",
               17: "nachname", 18: "vorname", 19: "strasse", 20: "hausnummer", 21: "plz", 22: "ort", 23: ("geburtsdatum", XmlOperation.open), 24: ("geburtsdatum", XmlOperation.close),
               26: "telefon", 27: "mobil", 28: "email", 29: "infoPerMail", 30: "sonstigeErreichbarkeit",
               34: "eintrittAktiv", 35: "endeAktiv",
               offsetLp + 1: "hl1", offsetLp + 2: "hl2", offsetLp + 3: "hl3", offsetLp + 4: "hl4", offsetLp + 5: "hl5",
               offsetLp + 6: "hl6", offsetLp + 7: "wa1", offsetLp + 8: "wa2", offsetLp + 9: "wa3", offsetLp + 10: "wa4",
               offsetLp + 11: "wa5", offsetLp + 12: "wa6",
               offsetAusbildung + 1: "ausbildungGA", offsetAusbildung + 2: "ausbildungTM",
               offsetAusbildung + 3: "ausbildungGF", offsetAusbildung + 4: "ausbildungZF",
               offsetAusbildung + 5: "ausbildungVF", offsetAusbildung + 6: "ausbildungFunk",
               offsetAusbildung + 7: "ausbildungMA", offsetAusbildung + 8: "ausbildungAT",
               offsetVerfuegbarkeit + 1: "verfWocheTag", offsetVerfuegbarkeit + 3: "verfWocheNacht",
               offsetVerfuegbarkeit + 5: "verfWochenendeTag", offsetVerfuegbarkeit + 7: "verWochenendeNacht"}

        idsAsList = ids.items()

        xml = "    <item>\n"

        operation = XmlOperation.openAndClose

        for pair in idsAsList:
            index = pair[0]

            if index >= len(row):
                continue

            if isinstance(pair[1], str):
                name = pair[1]
            else:
                if isinstance(pair[1], tuple):
                    name = pair[1][0]
                    operation = pair[1][1]

            cell = row[index]

            baseValue = cell.value

            value = self.__convertCellValue(baseValue, cell.ctype)

            if operation == XmlOperation.open:
                xml += "        <{0}>{1}".format(name, value)
            elif operation == XmlOperation.close:
                xml += "{1}</{0}>\n".format(name, value)
                operation = XmlOperation.openAndClose
            else:
                xml += "        <{0}>{1}</{0}>\n".format(name, value)
                operation = XmlOperation.openAndClose

        xml += "    </item>"
        return xml

    def __convertCellValue(self, baseValue, type):

        if type in self.mapCellType:
            result = self.mapCellType[type](baseValue)
        else:
            result = str(baseValue)

        return result

    def parse(self):

        book = xlrd.open_workbook(self.filename)
        sheet = book.sheet_by_name(self.sheetName)

        rows = self.__getRelevantRows(sheet)
        xml = self.__convertRowsToXml(rows)

        return xml

