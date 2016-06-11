###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...worksheet import Worksheet


class TestSharedStringCols(unittest.TestCase):
    """
    Test the _get_shared_string_cols Worksheet method for different col ranges.

    """

    def setUp(self):
        self.worksheet = Worksheet()

    def test_shared_string_cols_0(self):
        """Test Worksheet _get_shared_string_cols()"""

        exp = set([0])
        got = self.worksheet._get_shared_string_cols('A')

        self.assertEqual(got, exp)

    def test_shared_string_cols_1(self):
        """Test Worksheet _get_shared_string_cols()"""

        exp = set([0])
        got = self.worksheet._get_shared_string_cols(0)

        self.assertEqual(got, exp)

    def test_shared_string_cols_2(self):
        """Test Worksheet _get_shared_string_cols()"""

        exp = set([0,1,2,3,4,5])
        got = self.worksheet._get_shared_string_cols('A:F')

        self.assertEqual(got, exp)

    def test_shared_string_cols_3(self):
        """Test Worksheet _get_shared_string_cols()"""

        exp = set([0,1,2,3,4,5])
        got = self.worksheet._get_shared_string_cols(range(0, 6))

        self.assertEqual(got, exp)

    def test_shared_string_cols_4(self):
        """Test Worksheet _get_shared_string_cols()"""

        exp = set([0,1,2,3,4,5,10,12,13])
        got = self.worksheet._get_shared_string_cols('A:F,K,M:N')

        self.assertEqual(got, exp)
