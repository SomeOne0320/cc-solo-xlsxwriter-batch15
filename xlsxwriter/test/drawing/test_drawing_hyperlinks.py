###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
import xml.etree.ElementTree as ET
from io import BytesIO
from zipfile import ZipFile

from xlsxwriter.image import Image
from xlsxwriter.url import Url
from xlsxwriter.workbook import Workbook


def _parse_relationships(xml_data):
    # Parse a .rels file into a dict keyed by relationship id.
    namespace = (
        "{http://schemas.openxmlformats.org/package/2006/relationships}"
    )
    root = ET.fromstring(xml_data)

    relationships = {}
    for element in root.findall(namespace + "Relationship"):
        rel_id = element.attrib["Id"]
        relationships[rel_id] = {
            "type": element.attrib["Type"].rsplit("/", 1)[-1],
            "target": element.attrib["Target"],
            "target_mode": element.attrib.get("TargetMode", ""),
        }

    return relationships


def _parse_hlink_clicks(xml_data):
    # Parse the hlinkClick elements of a drawing into a list of r:id values.
    root = ET.fromstring(xml_data)

    namespace = (
        "{http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships}"
    )
    return [
        element.attrib[namespace + "id"]
        for element in root.iter(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}"
            "hlinkClick"
        )
    ]


def _bmp_image():
    # A minimal valid in-memory BMP so the test doesn't depend on a file.
    row = b"\x00\x00\xff" + b"\x00" * (4 - 3)
    pixel_data = row * 2

    file_header = b"BM"
    file_header += (62 + len(pixel_data)).to_bytes(4, "little")
    file_header += b"\x00" * 4
    file_header += (62).to_bytes(4, "little")

    info_header = (40).to_bytes(4, "little")
    info_header += (2).to_bytes(4, "little", signed=True)
    info_header += (2).to_bytes(4, "little", signed=True)
    info_header += (1).to_bytes(2, "little")
    info_header += (24).to_bytes(2, "little")
    info_header += b"\x00" * 24

    return BytesIO(file_header + info_header + pixel_data)


class TestDrawingHyperlinks(unittest.TestCase):
    """
    Test hyperlink registration for image and textbox drawing objects by
    generating a workbook and following the drawing relationship references.

    """

    def setUp(self):
        self.xlsx_buffer = BytesIO()
        self.workbook = Workbook(self.xlsx_buffer)

    def _close_and_open(self):
        self.workbook.close()
        return ZipFile(BytesIO(self.xlsx_buffer.getvalue()))

    def test_drawing_hyperlink_targets(self):
        """Link references in each drawing resolve to registered targets."""

        worksheet = self.workbook.add_worksheet()
        other_worksheet = self.workbook.add_worksheet()

        url = Url("https://example.com/page#section")
        url.tip = "reused link"

        # The same Url object is used by an image and a textbox on different
        # sheets so each sheet must register its own relationship.
        image = Image(_bmp_image())
        image.url = url
        worksheet.insert_image("B2", image)

        other_worksheet.insert_textbox(
            "B2", "Text box", {"url": url}
        )

        # The same target is registered again so de-duplication keeps a single
        # relationship for it on the first drawing.
        duplicate = Image(_bmp_image())
        duplicate.url = Url("https://example.com/page#section")
        worksheet.insert_image("B8", duplicate)

        # An external link with a fragment, an internal workbook link and an
        # unlinked object all share the same sheet.
        worksheet.insert_image(
            "B14",
            _bmp_image(),
            {"url": "https://example.org/doc#anchor"},
        )
        worksheet.insert_image(
            "B20",
            _bmp_image(),
            {"url": "internal:Sheet2!A1"},
        )
        worksheet.insert_textbox("B26", "No link")

        xlsx_file = self._close_and_open()

        # Object links created via options keep the fragment; a directly
        # assigned Url keeps its pre-existing relationship target.
        expected = {
            "xl/drawings/_rels/drawing1.xml.rels": {
                "https://example.com/page": "External",
                "https://example.org/doc#anchor": "External",
                "#Sheet2!A1": "",
            },
            "xl/drawings/_rels/drawing2.xml.rels": {
                "https://example.com/page": "External",
            },
        }

        for rels_filename, targets in expected.items():
            relationships = _parse_relationships(
                xlsx_file.read(rels_filename)
            )

            hyperlink_targets = {
                rel["target"]: rel["target_mode"]
                for rel in relationships.values()
                if rel["type"] == "hyperlink"
            }

            self.assertEqual(targets, hyperlink_targets)

            # Every hlinkClick r:id in the drawing must point at a hyperlink.
            drawing_filename = rels_filename.replace(
                "_rels/", ""
            ).replace(".rels", "")
            link_ids = _parse_hlink_clicks(
                xlsx_file.read(drawing_filename)
            )

            self.assertTrue(link_ids)
            for rel_id in link_ids:
                self.assertIn(rel_id, relationships)
                self.assertEqual("hyperlink", relationships[rel_id]["type"])

            # Image and link relationship ids can't collide.
            image_ids = [
                rel_id
                for rel_id, rel in relationships.items()
                if rel["type"] == "image"
            ]
            self.assertFalse(set(link_ids) & set(image_ids))

    def test_relationship_ids_are_independent_per_sheet(self):
        """Each sheet allocates its own ids starting from rId1."""

        worksheet = self.workbook.add_worksheet()
        other_worksheet = self.workbook.add_worksheet()

        url = Url("https://example.org")

        worksheet.insert_image("B2", _bmp_image(), {"url": url})
        other_worksheet.insert_image("B2", _bmp_image(), {"url": url})

        xlsx_file = self._close_and_open()

        for index in (1, 2):
            relationships = _parse_relationships(
                xlsx_file.read(
                    f"xl/drawings/_rels/drawing{index}.xml.rels"
                )
            )

            hyperlink_ids = [
                rel_id
                for rel_id, rel in relationships.items()
                if rel["type"] == "hyperlink"
            ]
            self.assertEqual(["rId1"], hyperlink_ids)

    def test_shared_url_tracks_each_sheet_independently(self):
        """A shared Url object must not leak rel ids between drawings."""

        worksheet = self.workbook.add_worksheet()
        other_worksheet = self.workbook.add_worksheet()

        url = Url("https://example.org/reused")

        # Sheet 1: the link is rId1 and the image is rId2.
        first_image = Image(_bmp_image())
        first_image.url = url
        worksheet.insert_image("B2", first_image)

        # Sheet 2: an unlinked image takes rId1 so the shared link is rId2.
        other_worksheet.insert_image("B2", _bmp_image())
        other_worksheet.insert_textbox("B8", "Box", {"url": url})

        xlsx_file = self._close_and_open()

        for index, expected_link_id in ((1, "rId1"), (2, "rId2")):
            rels_filename = f"xl/drawings/_rels/drawing{index}.xml.rels"
            relationships = _parse_relationships(xlsx_file.read(rels_filename))

            drawing_filename = f"xl/drawings/drawing{index}.xml"
            link_ids = _parse_hlink_clicks(xlsx_file.read(drawing_filename))

            self.assertEqual([expected_link_id], link_ids)

            rel = relationships[expected_link_id]
            self.assertEqual("hyperlink", rel["type"])
            self.assertEqual("https://example.org/reused", rel["target"])

    def test_caller_url_object_is_unchanged(self):
        """Saving must not leave registration state on the caller's Url."""

        worksheet = self.workbook.add_worksheet()
        other_worksheet = self.workbook.add_worksheet()

        url = Url("https://example.com/page#section")
        url.tip = "kept tip"
        original = vars(url).copy()

        image = Image(_bmp_image())
        image.url = url

        worksheet.insert_image("B2", image)
        other_worksheet.insert_textbox("B2", "Text box", {"url": url})

        # The object can also back a cell hyperlink with its own rules.
        worksheet.write_url("D2", url)

        self._close_and_open()

        self.assertEqual(original, vars(url))
        self.assertEqual("kept tip", url.tip)
        self.assertFalse(hasattr(url, "_rel_index"))

    def test_same_target_is_registered_once_per_sheet(self):
        """Repeated targets share one relationship id per drawing."""

        worksheet = self.workbook.add_worksheet()

        # Distinct objects with the same target must share one relationship.
        worksheet.insert_image(
            "B2",
            _bmp_image(),
            {"url": "https://example.com/repeated"},
        )
        worksheet.insert_image(
            "B8",
            _bmp_image(),
            {"url": "https://example.com/repeated"},
        )
        worksheet.insert_textbox(
            "B14",
            "Box",
            {"url": "https://example.com/repeated"},
        )

        xlsx_file = self._close_and_open()

        relationships = _parse_relationships(
            xlsx_file.read("xl/drawings/_rels/drawing1.xml.rels")
        )

        hyperlink_ids = [
            rel_id
            for rel_id, rel in relationships.items()
            if rel["type"] == "hyperlink"
        ]

        self.assertEqual(["rId1"], hyperlink_ids)

        drawing = xlsx_file.read("xl/drawings/drawing1.xml")
        self.assertEqual(
            ["rId1", "rId1", "rId1"], _parse_hlink_clicks(drawing)
        )
