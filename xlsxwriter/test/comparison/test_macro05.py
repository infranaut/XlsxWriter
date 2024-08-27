###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("macro05.xlsx")


    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""
        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()
        workbook.add_custom_ui(self.vba_dir + 'customUI-01.xml', version=2006)
        workbook.add_custom_ui(self.vba_dir + 'customUI14-01.xml', version=2007)
        worksheet.write('A1', 'Test')
        workbook.close()

        self.assertExcelEqual()
