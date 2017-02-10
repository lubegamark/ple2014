import csv
import logging
import os
import re
import time
import traceback

import pandas as pd
import requests
import xlrd

import settings

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')


class ExcelConverter(object):
    """
    Conversion of multiple excel sheets to csv files
    Adapted from http://strife.pl/2014/12/converting-large-xls-xlsx-files-to-csv-using-python/
    """

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
        target = target_folder + os.path.splitext(base)[0] + '.csv'
        csv_file = open(target, 'w+')
        wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
        first_sheet = True
        for sheet_name in wb.sheet_names():
            try:
                print("Start converting: {}".format(sheet_name))
                start_time = time.time()
                sh = wb.sheet_by_name(sheet_name)

                if sheet_name == 'Kyegegwa':  # Kyegegwa has completely different data
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

                print("Finished converting: {} in {} seconds".format(sheet_name, time.time() - start_time))

            except Exception as e:
                logging.error(str(e) + " " + traceback.format_exc())
            first_sheet = False
        csv_file.close()
        print("Overall Finished in {} seconds".format(time.time() - overall_start_time))

    @staticmethod
    def excel_to_csv_multiple(xls_file, target_folder, wb=None):
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
                print("Start converting: {}".format(sheet_name))
                start_time = time.time()
                target = target_folder + sheet_name.upper() + '.csv'
                sh = wb.sheet_by_name(sheet_name)

                if sheet_name == 'Kyegegwa':  # Kyegegwa has completely different data
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

                print("Finished converting: {} in {} seconds".format(sheet_name, time.time() - start_time))

            except Exception as e:
                logging.error(str(e) + " " + traceback.format_exc())

        print("Overall Finished in {} seconds".format(time.time() - overall_start_time, ))


class PLEInfo(object):
    @classmethod
    def get_rows_columns(cls, file=settings.MAIN_FILE):
        """
        Method prints how many columns and rows each sheet has
        :param file: File which has the data
        :return None:
        """
        cls.get_columns(file)
        cls.get_rows(file)
        # work_book = xlrd.open_workbook(file)
        # sheet_names = work_book.sheet_names()
        # print("District | Rows | Columns ")
        #
        # for sheet_name in sheet_names:
        #     sheet = work_book.sheet_by_name(sheet_name)
        #     print("{0} | {1} | {2} ".format(sheet.name, sheet.ncols, sheet.nrows,))

    @staticmethod
    def get_columns(file):
        """
        Method prints how many columns each sheet has
        :param file: File which has the data
        :return None:
        """
        work_book = xlrd.open_workbook(file)
        sheet_names = work_book.sheet_names()
        print("District | Columns")

        for sheet_name in sheet_names:
            sheet = work_book.sheet_by_name(sheet_name)
            print("{0} | {1}".format(sheet.name, sheet.ncols, ))

    @staticmethod
    def get_rows(file):
        """
        Method prints how many rows each sheet has
        :param file: File which has the data
        :return None:
        """
        work_book = xlrd.open_workbook(file)
        sheet_names = work_book.sheet_names()
        print("District | Rows")

        for sheet_name in sheet_names:
            sheet = work_book.sheet_by_name(sheet_name)
            print("{0} | {1}".format(sheet.name, sheet.nrows, ))


def find_csv_shape(folder):
    d = {}
    for path, folders, files in os.walk(folder):
        for file in files:
            f = os.path.join(path, file)
            csv = pd.read_csv(f)
            if len(csv.columns) in d:
                d[len(csv.columns)] += 1
            else:
                d[len(csv.columns)] = 1
                print("{0}\n{1} - {2}".format(len(csv.columns), file, csv.columns))
    total_districts = sum(d.values())
    print("Sheets with column length {}".format(d))
    for key, value in d.items():
        percentage = (value / total_districts) * 100
        print("{} - {:.2f}%".format(key, percentage))


def remove_unnamed(folder, right_size):
    for path, folders, files in os.walk(folder):
        for file in files:
            f = os.path.join(path, file)
            old_csv = pd.read_csv(f)
            if len(old_csv.columns) != right_size:
                new_csv = old_csv[old_csv.columns[~old_csv.columns.str.contains('Unnamed:')]]
                new_csv.to_csv(f, quoting=csv.QUOTE_ALL, index=False)


def get_required_columns(folder,
                         columns=(
                                 'DISTRICT', 'SCHOOL', 'CANDIDATE NUMBER', 'M/F', 'ENG', 'SCI', 'SST', 'MAT', 'AGG',
                                 'DIV')):
    for dirpath, dirs, filesnames in os.walk(folder):
        for filename in filesnames:
            file = os.path.join(dirpath, filename)
            remove_empty_records(file, columns)


def strip_records(location):
    """
    Strip all string columns
    """
    if os.path.isfile(location):
        df = pd.read_csv(location)
        # Strip all String Columns
        df_obj = df.select_dtypes(['object'])
        df[df_obj.columns] = df_obj.apply(lambda x: x.str.strip())
        df.to_csv(location, quoting=csv.QUOTE_ALL, index=False)
    elif os.path.isdir(location):
        for path, folders, files in os.walk(location):
            for f in files:
                file = os.path.join(location, f)
                df = pd.read_csv(file)
                # Strip all String Columns
                df_obj = df.select_dtypes(['object'])
                df[df_obj.columns] = df_obj.apply(lambda x: x.str.strip())
                df.to_csv(file, quoting=csv.QUOTE_ALL, index=False)


def remove_empty_records(location):
    """
    Remove all records empty records
    """
    if os.path.isfile(location):
        df = pd.read_csv(location)
        # Drop all records without data
        df.dropna(how='all', inplace=True)
        df.to_csv(location, quoting=csv.QUOTE_ALL, index=False)
    elif os.path.isdir(location):
        for path, folders, files in os.walk(location):
            for f in files:
                file = os.path.join(location, f)
                df = pd.read_csv(file)
                # Drop all records without data
                df.dropna(how='all', inplace=True)
                df.to_csv(file, quoting=csv.QUOTE_ALL, index=False)


def correct_headers(location):
    """
    Some files have inconsistent headings.
    These are corrected her
    """
    if os.path.isfile(location):
        df = pd.read_csv(location)
        # Special Case for Kisoro, it's district column name is kisoro
        df.rename(columns={'F/M': 'M/F', 'SCIE': 'SCI', 'MATH': 'MAT', 'CNDIDATE NUMBER': 'CANDIDATE NUMBER',
                           'KISORO': 'DISTRICT'},
                  inplace=True)
        df.columns = df.columns.str.strip()
        df.columns = df.columns.str.upper()
        df.to_csv(location, quoting=csv.QUOTE_ALL, index=False)
    elif os.path.isdir(location):
        for path, folders, files in os.walk(location):
            for f in files:
                file = os.path.join(location, f)
                df = pd.read_csv(file)
                # Special Case for Kisoro, it's district column name is kisoro
                df.rename(columns={'F/M': 'M/F', 'SCIE': 'SCI', 'MATH': 'MAT', 'CNDIDATE NUMBER': 'CANDIDATE NUMBER',
                                   'KISORO': 'DISTRICT'},
                          inplace=True)
                df.columns = df.columns.str.strip()
                df.columns = df.columns.str.upper()
                df.to_csv(file, quoting=csv.QUOTE_ALL, index=False)


def convert_numeric_values(location):
    """
    Convert Numeric fields to numeric data types
    """
    if os.path.isfile(location):
        df = pd.read_csv(location)
        df['DIV'].replace('U', '0', inplace=True)
        df[['MAT', 'SCI', 'SST', 'ENG', 'AGG', 'DIV']] = df[
            ['MAT', 'SCI', 'SST', 'ENG', 'AGG', 'DIV']].apply(pd.to_numeric, errors='coerce')
        df.to_csv(location, quoting=csv.QUOTE_ALL, index=False)
    elif os.path.isdir(location):
        for path, folders, files in os.walk(location):
            for f in files:
                file = os.path.join(location, f)
                df = pd.read_csv(file)
                df['DIV'].replace("U", "0", inplace=True)
                df[['MAT', 'SCI', 'SST', 'ENG', 'AGG', 'DIV']] = df[
                    ['MAT', 'SCI', 'SST', 'ENG', 'AGG', 'DIV']].apply(pd.to_numeric, errors='coerce')
                df.to_csv(file, quoting=csv.QUOTE_ALL, index=False)


def download_file(url):
    """
    Got from http://stackoverflow.com/a/16696317/5117592
    """
    local_filename = url.split('/')[-1]
    file = os.path.join('data/original', local_filename)
    # NOTE the stream=True parameter
    r = requests.get(url, stream=True)
    with open(file, 'wb') as f:
        for chunk in r.iter_content(chunk_size=1024):
            if chunk:  # filter out keep-alive new chunks
                f.write(chunk)
                # f.flush() commented by recommendation from J.F.Sebastian
    return file


def download_ple():
    if not os.path.exists(settings.MAIN_FILE):
        download_file(
            'http://ugandajournalistsresourcecentre.com/wp-content/uploads/2015/05/PLE-Results-2014.ALL-CANDIDATES.xlsx'
        )
    return
