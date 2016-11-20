import argparse
import os
import sys

from ExcelReader import ExcelReader


def usage():
    sys.stderr.write("Usage: %s [-e] [-o outputfile] [inputfile]\n" % sys.argv[0])


def isValidFile(parser, arg):
    if not os.path.isfile(arg):
        parser.error("The file %s does not exist!" % arg)
    else:
        return arg


if __name__ == '__main__':
    """
    Pass in the name of the incoming file and the
    phrase as command line arguments. Use sys.argv[]
    """

    parser = argparse.ArgumentParser(description='Converts the Feuerwehr Mitglieder_aktuell.xmlx to an XML file')
    parser.add_argument('-i', '--input', required=True, help='input Office Open XML *.xmlx file', metavar='FILE', type=lambda x: isValidFile(parser, x))
    parser.add_argument('-o', '--output', required=True, help='output filename of the XML target', metavar='FILE')

    args = parser.parse_args()

    input = os.path.abspath(os.path.expanduser(args.input))
    output = os.path.abspath(os.path.expanduser(args.output))

    odfReader = ExcelReader(input)
    xml = odfReader.parse()

    with open(output, "w", encoding="utf-8") as f:
        f.write(xml)
