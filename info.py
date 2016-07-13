"""
Simple file to get some general details about our data
"""
import xlrd
import settings


class PLEInfo(object):

    @staticmethod
    def get_rows_columns(file=settings.MAIN_FILE):
        """
        Method prints how many columns and rows each sheet has
        :param file: File which has the data
        :return None:
        """
        work_book = xlrd.open_workbook(file)
        sheet_names = work_book.sheet_names()
        print("District | Rows | Columns ")

        for sheet_name in sheet_names:
            sheet = work_book.sheet_by_name(sheet_name)
            print("{0} | {1} | {2} ".format(sheet.name, sheet.ncols,sheet.nrows,))
