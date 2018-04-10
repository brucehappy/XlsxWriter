###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'chart_data_labels21.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'pie'})

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],

        ]

        worksheet.write_column('A1', data[0])
        worksheet.write_column('B1', data[1])
        worksheet.write_column('C1', data[2])

        chart.add_series({
            'values': '=Sheet1!$A$1:$A$5',
            'data_labels': {
                'value': True,
                'category': True,
                'series_name': True,
                'percentage': True,
                'separator': ';',
                'leader_lines': True,
                'position': 'inside_end',
                'legend_key': True,
                'num_format': '#,##0.00',
                'font': {'name': 'Consolas', 'baseline': 1 * -1, 'pitch_family': 49, 'charset': 0}
            },
        })

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
