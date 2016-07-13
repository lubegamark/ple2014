"""
Conversion of multiple excel sheets to csv files
Adapted from http://strife.pl/2014/12/converting-large-xls-xlsx-files-to-csv-using-python/
"""

import csv
import logging
import os
import re
import time
import traceback

import xlrd

import settings

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')


class ExcelConverter(object):
    def __init__(self):
        pass

    @staticmethod
    def excel_to_csv(wb=None, xls_file=settings.MAIN_FILE, target_folder=settings.PROCESSED_FOLDER):
        """
        Convert an excel file(.xls/.xslx) to CSV writing all sheets to one file


        :param xls_file: Original excel file to be converted
        :param target_folder: Folder in which the generated file will be stored
        :param wb: Loaded workbook
        """
        overall_start_time = time.time()
        print("Start converting")
        if wb:
            print("Workbook provided")

        else:
            print("No workbook provided")
            wb = xlrd.open_workbook(xls_file)

        base = os.path.basename(xls_file)
        target = target_folder+os.path.splitext(base)[0]+'.csv'
        csv_file = open(target, 'w+')
        wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
        first_sheet = True
        for sheet_name in wb.sheet_names():
            try:
                print("Start converting: %s" % sheet_name)
                start_time = time.time()
                sh = wb.sheet_by_name(sheet_name)

                if sheet_name == 'Kyegegwa':#Kyegegwa has completely different data
                    print("Moving on, Kyegegwa has different data")
                    continue

                if first_sheet:
                    range_start = 0
                else:
                    range_start = 1

                for row in range(range_start, sh.nrows):
                    row_values = sh.row_values(row)

                    new_values = []
                    for s in row_values:
                        str_value = (str(s).strip())

                        is_int = bool(re.match("^([0-9]+)\.0$", str_value))

                        if is_int:
                            str_value = int(float(str_value))
                        else:
                            is_float = bool(re.match("^([0-9]+)\.([0-9]+)$", str_value))
                            is_long = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", str_value))

                            if is_float:
                                str_value = float(str_value)

                            if is_long:
                                str_value = int(float(str_value))

                        new_values.append(str_value)

                    wr.writerow(new_values)

                print("Finished converting: %s in %s seconds" % (sheet_name, time.time() - start_time))

            except Exception as e:
                logging.error(str(e) + " " + traceback.format_exc())
            first_sheet = False
        csv_file.close()
        print("Overall Finished in %s seconds", time.time() - overall_start_time)

    @staticmethod
    def excel_to_csv_multiple(xls_file=settings.MAIN_FILE, target_folder=settings.PROCESSED_FOLDER, wb=None):
        """
        Convert an excel file(.xls/.xslx) to CSV writing each sheet to a separate file

        :param xls_file: Original excel file to be converted
        :param target_folder: Folder in which the generated files will be stored
        """
        overall_start_time = time.time()
        print("Start converting")
        if not wb:
            wb = xlrd.open_workbook(xls_file)
        for sheet_name in wb.sheet_names():
            try:
                print("Start converting: %s" % sheet_name)
                start_time = time.time()
                target = target_folder+sheet_name.upper()+'.csv'
                sh = wb.sheet_by_name(sheet_name)

                if sheet_name == 'Kyegegwa':#Kyegegwa has completely different data
                    print("Moving on, Kyegegwa has different data")
                    continue

                csv_file = open(target, 'w')
                wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)

                for row in range(sh.nrows):
                    row_values = sh.row_values(row)

                    new_values = []
                    for s in row_values:
                        str_value = (str(s))

                        is_int = bool(re.match("^([0-9]+)\.0$", str_value))

                        if is_int:
                            str_value = int(float(str_value))
                        else:
                            is_float = bool(re.match("^([0-9]+)\.([0-9]+)$", str_value))
                            is_long = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", str_value))

                            if is_float:
                                str_value = float(str_value)

                            if is_long:
                                str_value = int(float(str_value))

                        new_values.append(str_value)

                    wr.writerow(new_values)

                csv_file.close()

                print("Finished converting: %s in %s seconds" % (sheet_name, time.time() - start_time))

            except Exception as e:
                logging.error(str(e) + " " + traceback.format_exc())

        print("Overall Finished in %s seconds", time.time() - overall_start_time)