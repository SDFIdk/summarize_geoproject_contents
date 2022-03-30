# -*- coding: utf-8 -*-
"""
A wrapper for https://xlsxwriter.readthedocs.io/, with utilities to write xlsx.
No, it isn't pretty, but its just for import here...

Created by: Hanne L. Petersen <halpe@sdfe.dk>
Created on: 5 Feb 2018
"""


class OfficeUtils(object):

    @staticmethod
    def multi_lists2xlsx(lstlst, xlsx_path, headings=None, config=[]):
        """Write a list of lists to xlsx file."""
        if len(lstlst) < 1:
            # TODO: print/set_msg
            print("No data, no file written")
            return

        import xlsxwriter
        workbook = xlsxwriter.Workbook(xlsx_path)

        # Add a bold format to use to highlight cells
        bold_fmt = workbook.add_format({'bold': True})
        # Add locked formats for protecting cells
        unlocked_fmt = workbook.add_format({'locked': False})
        locked_fmt = workbook.add_format({'locked': True})
        formats = {'bold': bold_fmt, 'unlocked': unlocked_fmt, 'locked': locked_fmt}

        for n, l in enumerate(lstlst):  # yes, enumerate() is py2 and py3
            if 'sheetname' in config:
                sheetname = config['sheetname'][n]
            else:
                sheetname = 'Sheet{}'.format(n)

            worksheet = workbook.add_worksheet(sheetname)
            OfficeUtils.write_sheet(worksheet, l, headings, config, formats)

        workbook.close()

    @staticmethod
    def lists2xlsx(lstlst, xlsx_path, headings=None, config=[]):
        """Write a list of lists to xlsx file."""
        if len(lstlst) < 1:
            # TODO: print/set_msg
            print("No data, no file written")
            return

        import xlsxwriter
        if 'sheetname' in config:
            sheetname = config['sheetname']
        else:
            sheetname = 'Sheet1'
        workbook = xlsxwriter.Workbook(xlsx_path)

        # Add a bold format to use to highlight cells
        bold_fmt = workbook.add_format({'bold': True})
        # Add locked formats for protecting cells
        unlocked_fmt = workbook.add_format({'locked': False})
        locked_fmt = workbook.add_format({'locked': True})

        worksheet = workbook.add_worksheet(sheetname)

        formats = {'bold': bold_fmt, 'unlocked': unlocked_fmt, 'locked': locked_fmt}
        OfficeUtils.write_sheet(worksheet, lstlst, headings, config, formats)

        workbook.close()

    @staticmethod
    def write_sheet(worksheet, lstlst, headings, config, formats):
        row_counter = 0

        # Default column widths
        worksheet.set_column(0, 0, 50)  # first column wider
        worksheet.set_column(1, len(lstlst[0])-1, 30)

        # Custom column widths
        if 'widths' in config:
            for (c, w) in enumerate(config['widths']):
                worksheet.set_column(c, c, w)  # first_col, last_col, width

        # Freeze panes
        if 'freeze' in config and config['freeze']:
            worksheet.freeze_panes(1, 0)  # Freeze top row

        # Lock/protect columns
        if 'protect_cols' in config:
            # https://stackoverflow.com/a/40891864/7121793
            worksheet.set_column('A:XDF', None, formats['unlocked'])  # Unlock all
            worksheet.protect()
            for col in config['protect_cols']:
                worksheet.set_column(col+':'+col, None, formats['locked'])

        # Column headings
        if 'headings' in config:
            worksheet.write_row(row_counter, 0, config['headings'], formats['bold'])
            row_counter += 1
        if headings:  # The old way for backwards compatibility...
            worksheet.write_row(row_counter, 0, headings, formats['bold'])
            row_counter += 1

        for r in lstlst:
            worksheet.write_row(row_counter, 0, r)
            row_counter += 1
