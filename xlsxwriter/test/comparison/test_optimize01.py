###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2019, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('optimize01.xlsx')

    def _test_create_file(self, extra_options={}):
        """Test the creation of a simple XlsxWriter file."""

        options = {
            'constant_memory': True,
            'strings_to_numbers': True,
            'in_memory': False
        }
        options.update(extra_options)
        workbook = Workbook(self.got_filename, options)
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'Hello')
        worksheet.write('A2', '123')

        workbook.close()

        self.assertExcelEqual()

    def test_create_file(self):
        self._test_create_file()

    def test_create_file_with_buffer(self):
        # This tests no rows being written until close
        self._test_create_file({'constant_memory_row_buffer': 1000})
