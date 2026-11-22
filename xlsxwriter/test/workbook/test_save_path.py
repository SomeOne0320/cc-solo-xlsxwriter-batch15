###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import os
import re
import tempfile
import unittest
from io import BytesIO
from zipfile import ZipFile

from xlsxwriter.packager import Packager
from xlsxwriter.workbook import Workbook


class LegacyPackager(Packager):
    """
    A Packager that ignores the direct-to-Zip stream and uses the original
    temp file/in-memory protocol. Used to check that both save paths produce
    the same package.

    """

    def _set_zip_file(self, zip_file):
        # Deliberately ignore the Zip container to exercise the fallback.
        pass


class LegacyWorkbook(Workbook):
    """
    A Workbook with the legacy packager, i.e., the pre-existing save path.

    """

    def _get_packager(self):
        return LegacyPackager()


def _normalize(data):
    # Remove document timestamps so that workbooks created at different
    # times can be compared.
    return re.sub(rb"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z", b"<TIMESTAMP>", data)


class TestSavePath(unittest.TestCase):
    """
    Test the workbook save/packaging path.

    """

    def _add_test_data(self, workbook, in_memory=False):
        # Add a small but varied multi-sheet workbook.
        image_dir = "xlsxwriter/test/comparison/images/"
        bold = workbook.add_format({"bold": True})

        for sheet_num in range(3):
            worksheet = workbook.add_worksheet(f"Sheet{sheet_num + 1}")
            worksheet.write_string(0, 0, "Hello", bold)
            worksheet.write_number(1, 0, 123)
            worksheet.write_formula(2, 0, "=SUM(A2)", bold)
            worksheet.write("A4", "café")
            worksheet.insert_image(5, 0, image_dir + "red.png")

        return workbook

    def _compare(self, got_file, exp_file):
        got_zip = ZipFile(got_file)
        exp_zip = ZipFile(exp_file)

        self.assertEqual(sorted(got_zip.namelist()), sorted(exp_zip.namelist()))

        for filename in exp_zip.namelist():
            self.assertEqual(
                _normalize(got_zip.read(filename)),
                _normalize(exp_zip.read(filename)),
                f"Different data for {filename}",
            )

        # The Zip metadata, such as timestamps and compression, should match.
        got_info = [
            (z.filename, z.date_time, z.compress_type) for z in got_zip.infolist()
        ]
        exp_info = [
            (z.filename, z.date_time, z.compress_type) for z in exp_zip.infolist()
        ]
        self.assertEqual(got_info, exp_info)

    def test_save_path_matches_legacy_temp_files(self):
        """Test that streaming to the Zip matches the temp file save path."""

        tmpdir = tempfile.mkdtemp()
        got_file = os.path.join(tmpdir, "streamed.xlsx")
        exp_file = os.path.join(tmpdir, "legacy.xlsx")

        self._add_test_data(Workbook(got_file)).close()
        self._add_test_data(LegacyWorkbook(exp_file)).close()

        self._compare(got_file, exp_file)

    def test_save_path_matches_legacy_in_memory(self):
        """Test the in_memory save path against the legacy path."""

        tmpdir = tempfile.mkdtemp()
        got_file = os.path.join(tmpdir, "streamed.xlsx")
        exp_file = os.path.join(tmpdir, "legacy.xlsx")

        self._add_test_data(Workbook(got_file, {"in_memory": True})).close()
        self._add_test_data(
            LegacyWorkbook(exp_file, {"in_memory": True})
        ).close()

        self._compare(got_file, exp_file)

    def test_save_path_matches_legacy_bytes_io(self):
        """Test saving to a caller provided byte stream."""

        got_stream = BytesIO()
        exp_file = os.path.join(tempfile.mkdtemp(), "legacy.xlsx")

        self._add_test_data(Workbook(got_stream)).close()
        self._add_test_data(LegacyWorkbook(exp_file)).close()

        got_file = os.path.join(tempfile.mkdtemp(), "streamed.xlsx")
        with open(got_file, "wb") as fh:
            fh.write(got_stream.getvalue())

        self._compare(got_file, exp_file)

    def test_save_path_avoids_intermediate_files(self):
        """Test that the normal save path doesn't create temp files."""

        real_mkstemp = tempfile.mkstemp
        created = []

        def counting_mkstemp(*args, **kwargs):
            created.append(1)
            return real_mkstemp(*args, **kwargs)

        filename = os.path.join(tempfile.mkdtemp(), "test.xlsx")
        workbook = self._add_test_data(Workbook(filename))

        tempfile.mkstemp = counting_mkstemp
        try:
            workbook.close()
        finally:
            tempfile.mkstemp = real_mkstemp

        self.assertEqual(created, [])

        # The output should still be a valid, complete package.
        with ZipFile(filename) as zip_file:
            self.assertIsNone(zip_file.testzip())
            self.assertIn("xl/worksheets/sheet1.xml", zip_file.namelist())
            self.assertIn("xl/media/image1.png", zip_file.namelist())

    def test_save_path_constant_memory(self):
        """Test that constant_memory mode still writes its rows correctly."""

        tmpdir = tempfile.mkdtemp()
        got_file = os.path.join(tmpdir, "constant.xlsx")
        exp_file = os.path.join(tmpdir, "legacy.xlsx")

        for filename, workbook_class in (
            (got_file, Workbook),
            (exp_file, LegacyWorkbook),
        ):
            workbook = workbook_class(filename, {"constant_memory": True})
            worksheet = workbook.add_worksheet()
            for row in range(100):
                worksheet.write_row(row, 0, [row, "text", row + 0.5])
            workbook.close()

        self._compare(got_file, exp_file)
