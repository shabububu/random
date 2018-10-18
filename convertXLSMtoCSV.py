#!/usr/bin/env python

# Based off of: https://stackoverflow.com/questions/41230716/convert-xlsm-to-csv

import csv
import xlrd
from argparse import ArgumentParser
import os.path

parser = ArgumentParser()
parser.add_argument("-f", "--file", dest="xlsm_filename",
                    help="path to *.xlsm file", required=True)
args = parser.parse_args()
print args

if not os.path.isfile(args.xlsm_filename):
  print("The file %s is not a valid file.!" % args.xlsm_filename)
  exit()

workbook = xlrd.open_workbook(args.xlsm_filename)
for sheet in workbook.sheets():
    with open('{}.csv'.format(sheet.name), 'wb') as f:
        writer = csv.writer(f)
        print("Sheet name: %s, # Rows: %d" % (sheet.name,sheet.nrows))
        writer.writerows([x.encode('utf-8') for x in sheet.row_values(row)] for row in range(sheet.nrows))
