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
import zipfile

from xlsxwriter.workbook import Workbook


def _chart_xml(path):
    # Return the chart XML parts contained in an xlsx file.
    with zipfile.ZipFile(path) as xlsx:
        names = sorted(
            name
            for name in xlsx.namelist()
            if re.match(r"xl/charts/chart\d+\.xml$", name)
        )
        return [xlsx.read(name).decode("utf-8") for name in names]


def _cache_blocks(xml):
    # Extract the numCache/strCache blocks from a chart XML string.
    return re.findall(
        r"<c:(?:num|str)Cache>.*?</c:(?:num|str)Cache>", xml, re.DOTALL
    )


def _point_values(xml, formula):
    # Return the cached (index, value) points for a given range formula.
    match = re.search(
        re.escape(f"<c:f>{formula}</c:f>")
        + r"<c:(?:num|str)Cache>(.*?)</c:(?:num|str)Cache>",
        xml,
        re.DOTALL,
    )
    return re.findall(
        r'<c:pt idx="(\d+)"><c:v>([^<]*)</c:v></c:pt>', match.group(1)
    )


class TestChartCachedData(unittest.TestCase):
    """
    Test that chart numCache/strCache preview data is generated when a
    workbook is written in constant_memory mode.

    """

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def _path(self, name):
        return os.path.join(self.tmpdir, name)

    def _build_workbook(self, path, constant_memory):
        # Build the same workbook in either storage mode. The data is written
        # row by row, as required by constant_memory mode.
        workbook = Workbook(path, {"constant_memory": constant_memory})
        worksheet = workbook.add_worksheet()

        rows = [
            (1, 10, "alpha"),
            (2, 20, "beta"),
            (None, None, "gamma"),
            (4, 40, "delta"),
            (5, 50, "epsilon"),
        ]

        for row, (a_val, b_val, c_val) in enumerate(rows):
            if a_val is not None:
                worksheet.write(row, 0, a_val)
            if row == 2:
                # A formula with a pre-calculated result.
                worksheet.write_formula(row, 1, "=1+1", None, 2)
            elif b_val is not None:
                worksheet.write(row, 1, b_val)
            worksheet.write(row, 2, c_val)

        chart = workbook.add_chart({"type": "line"})
        chart.add_series(
            {
                "values": "=Sheet1!$A$1:$A$5",
                "categories": "=Sheet1!$C$1:$C$5",
            }
        )
        chart.add_series({"values": "=Sheet1!$B$1:$B$5"})
        worksheet.insert_chart("E2", chart)

        workbook.close()

    def test_chart_cache_present_in_constant_memory(self):
        """Chart preview data must be written in constant_memory mode."""

        path = self._path("optimized.xlsx")
        self._build_workbook(path, True)
        xml = _chart_xml(path)[0]

        self.assertIn("<c:numCache>", xml)
        self.assertIn("<c:strCache>", xml)

    def test_chart_cache_matches_normal_mode(self):
        """Chart XML (including caches) must match the normal mode output."""

        normal = self._path("normal.xlsx")
        optimized = self._path("optimized.xlsx")

        self._build_workbook(normal, False)
        self._build_workbook(optimized, True)

        # The chart part is byte for byte identical across the two modes.
        self.assertEqual(_chart_xml(normal), _chart_xml(optimized))

    def test_gap_keeps_later_values_in_position(self):
        """A missing cell leaves a gap without shifting later points."""

        path = self._path("gap.xlsx")
        self._build_workbook(path, True)
        xml = _chart_xml(path)[0]

        points = _point_values(xml, "Sheet1!$A$1:$A$5")
        self.assertEqual(
            points,
            [("0", "1"), ("1", "2"), ("3", "4"), ("4", "5")],
        )

    def test_formula_uses_stored_result(self):
        """A formula point uses the saved pre-calculated result."""

        path = self._path("formula.xlsx")
        self._build_workbook(path, True)
        xml = _chart_xml(path)[0]

        points = _point_values(xml, "Sheet1!$B$1:$B$5")
        self.assertIn(("2", "2"), points)

    def test_horizontal_range(self):
        """A range taken across a single row is cached column by column."""

        path = self._path("horizontal.xlsx")
        workbook = Workbook(path, {"constant_memory": True})
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, 1)
        worksheet.write(0, 1, 2)
        # Column C (index 2) is left blank on purpose.
        worksheet.write(0, 3, 4)
        worksheet.write(0, 4, 5)
        # Additional rows so the first row is flushed.
        worksheet.write(1, 0, 9)
        worksheet.write(2, 0, 9)

        chart = workbook.add_chart({"type": "line"})
        chart.add_series({"values": "=Sheet1!$A$1:$E$1"})
        worksheet.insert_chart("G1", chart)
        workbook.close()

        xml = _chart_xml(path)[0]
        points = _point_values(xml, "Sheet1!$A$1:$E$1")
        self.assertEqual(
            points,
            [("0", "1"), ("1", "2"), ("3", "4"), ("4", "5")],
        )

    def test_multi_segment_range_concatenated_in_order(self):
        """Separate segments of a series are joined in reference order."""

        path = self._path("multi.xlsx")
        workbook = Workbook(path, {"constant_memory": True})
        worksheet = workbook.add_worksheet()

        for row, value in enumerate([10, 20, 30, 40, 50, 60]):
            worksheet.write(row, 0, value)

        formula = "(Sheet1!$A$4:$A$6,Sheet1!$A$1:$A$2)"
        chart = workbook.add_chart({"type": "line"})
        chart.add_series({"values": "=" + formula})
        worksheet.insert_chart("C1", chart)
        workbook.close()

        xml = _chart_xml(path)[0]
        self.assertIn(f"<c:f>{formula}</c:f>", xml)

        match = re.search(
            re.escape(f"<c:f>{formula}</c:f>")
            + r"<c:numCache>(.*?)</c:numCache>",
            xml,
            re.DOTALL,
        )
        points = re.findall(r'<c:pt idx="(\d+)"><c:v>([^<]*)</c:v></c:pt>',
                            match.group(1))
        self.assertEqual(
            [value for _, value in points],
            ["40", "50", "60", "10", "20"],
        )

    def test_quoted_sheet_name_with_spaces(self):
        """Charts can reference a sheet name containing spaces/quotes."""

        path = self._path("quoted.xlsx")
        workbook = Workbook(path, {"constant_memory": True})
        worksheet = workbook.add_worksheet("My Data")

        for row, value in enumerate([1, 2, 3]):
            worksheet.write(row, 0, value)

        chart = workbook.add_chart({"type": "line"})
        chart.add_series({"values": "='My Data'!$A$1:$A$3"})
        worksheet.insert_chart("C1", chart)
        workbook.close()

        xml = _chart_xml(path)[0]
        self.assertIn("<c:numCache>", xml)
        self.assertIn("<c:f>'My Data'!$A$1:$A$3</c:f>", xml)
        self.assertEqual(
            _point_values(xml, "'My Data'!$A$1:$A$3"),
            [("0", "1"), ("1", "2"), ("2", "3")],
        )

    def test_explicit_preview_data_is_preserved(self):
        """User supplied values_data takes precedence over cell data."""

        path = self._path("user_data.xlsx")
        workbook = Workbook(path, {"constant_memory": True})
        worksheet = workbook.add_worksheet()

        for row, value in enumerate([1, 2, 3]):
            worksheet.write(row, 0, value)

        chart = workbook.add_chart({"type": "line"})
        chart.add_series(
            {
                "values": "=Sheet1!$A$1:$A$3",
                "values_data": [7, 8, 9],
            }
        )
        worksheet.insert_chart("C1", chart)
        workbook.close()

        xml = _chart_xml(path)[0]
        points = _point_values(xml, "Sheet1!$A$1:$A$3")
        self.assertEqual(
            [value for _, value in points], ["7", "8", "9"]
        )

    def test_temp_chart_data_file_is_removed(self):
        """The chart data temp file must not be left behind after close()."""

        path = self._path("cleanup.xlsx")
        workbook = Workbook(path, {"constant_memory": True})
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, 1)

        chart = workbook.add_chart({"type": "line"})
        chart.add_series({"values": "=Sheet1!$A$1"})
        worksheet.insert_chart("C1", chart)
        workbook.close()

        self.assertIsNone(worksheet.chart_data_filename)
        self.assertIsNone(worksheet.chart_data_fh)

    def test_cache_blocks_are_parity_with_normal_mode(self):
        """The cache blocks themselves match the normal mode workbook."""

        normal = self._path("normal_blocks.xlsx")
        optimized = self._path("optimized_blocks.xlsx")

        self._build_workbook(normal, False)
        self._build_workbook(optimized, True)

        normal_blocks = _cache_blocks(_chart_xml(normal)[0])
        optimized_blocks = _cache_blocks(_chart_xml(optimized)[0])

        self.assertEqual(normal_blocks, optimized_blocks)

    def test_uncharted_cell_types_match_normal_mode(self):
        """Boolean cells add no point while blank cells do, matching normal."""

        normal = self._path("normal_types.xlsx")
        optimized = self._path("optimized_types.xlsx")

        def build(path, constant_memory):
            workbook = Workbook(path, {"constant_memory": constant_memory})
            worksheet = workbook.add_worksheet()
            cell_format = workbook.add_format({"bold": True})

            worksheet.write_boolean(0, 0, True)  # uncharted: no point
            worksheet.write_number(1, 0, 42)
            worksheet.write_blank(2, 0, None, cell_format)  # -> ""
            worksheet.write_formula(3, 0, "=1+1", None, 2)
            worksheet.write_number(4, 0, 7)

            chart = workbook.add_chart({"type": "line"})
            chart.add_series({"values": "=Sheet1!$A$1:$A$5"})
            worksheet.insert_chart("C1", chart)
            workbook.close()

        build(normal, False)
        build(optimized, True)

        # The full chart part must be identical across the two modes.
        self.assertEqual(_chart_xml(normal), _chart_xml(optimized))


if __name__ == "__main__":
    unittest.main()
