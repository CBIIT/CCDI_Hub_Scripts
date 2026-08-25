#!/usr/bin/env python3
"""Tests for parsing the remote site announcement Markdown format."""

import importlib.util
from pathlib import Path
import unittest


MODULE_PATH = Path(__file__).with_name("yaml_to_pdf_generator.py")
SPEC = importlib.util.spec_from_file_location("release_notes_generator", MODULE_PATH)
release_notes_generator = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(release_notes_generator)


SAMPLE_MARKDOWN = """\
# First release
### August 31, 2026 | Release Notes

Intro with **bold text** and a [link](https://example.org).

### Updates

- First item
  - Nested item

| Property | Value |
| --- | --- |
| id | catalog_release_08312026 |
| version | v1.5.10 |
| slug | A useful summary |
| contentType | Clinical,Imaging |

# Earlier release
### June 10, 2026 | Release Notes

Earlier content.

| Property | Value |
| --- | --- |
| version | v1.5.9 |
"""

SUMMARY_MARKDOWN = """\
# A late summer bloom of pediatric data resources
### August 31, 2026 | Release Notes

#### v1.5.10 Summary

- **36th** Release
- **541** Datasets
- **164** Resources
  - **25** Analytical Tools
  - **9** Biorepositories

| Property | Value |
| --- | --- |
| version | v1.5.10 |
"""

LOOSE_NESTED_LIST_MARKDOWN = """\
# New resources
### August 31, 2026 | Release Notes

### Data Updates

- African Cancer Registry Network

  - **New Dataset** - AFCRN Database
- Alaska Native Tumor Registry

  - **New Dataset** - Cancer in Alaska Native People

| Property | Value |
| --- | --- |
| version | v1.5.10 |
"""


class MarkdownParsingTests(unittest.TestCase):
    def setUp(self):
        self.generator = release_notes_generator.ReleaseNotesPDFGenerator()

    def test_parses_releases_and_metadata(self):
        releases = self.generator.parse_markdown_releases(SAMPLE_MARKDOWN)

        self.assertEqual(len(releases), 2)
        self.assertEqual(releases[0]["title"], "First release")
        self.assertEqual(releases[0]["date"], "August 31, 2026")
        self.assertEqual(releases[0]["version"], "v1.5.10")
        self.assertEqual(releases[0]["slug"], "A useful summary")
        self.assertEqual(releases[0]["dataType"], "Clinical,Imaging")
        self.assertEqual(releases[1]["version"], "v1.5.9")

    def test_converts_markdown_body_to_html_and_removes_property_table(self):
        release = self.generator.parse_markdown_releases(SAMPLE_MARKDOWN)[0]

        self.assertIn("<strong>bold text</strong>", release["fullText"])
        self.assertIn('<a href="https://example.org">link</a>', release["fullText"])
        self.assertIn("<ul>", release["fullText"])
        self.assertNotIn("Property", release["fullText"])
        self.assertNotIn("catalog_release_08312026", release["fullText"])

    def test_preserves_summary_heading_emphasis_and_list_levels(self):
        release = self.generator.parse_markdown_releases(SUMMARY_MARKDOWN)[0]
        elements = self.generator.parse_html_content(release["fullText"])

        self.assertEqual(elements[0].style.name, "SubsectionHeader")
        self.assertEqual(elements[0].text, "v1.5.10 Summary")

        list_items = elements[1:]
        self.assertEqual(
            [item.style.name for item in list_items],
            [
                "ListItem",
                "ListItem",
                "ListItem",
                "NestedListItem",
                "NestedListItem",
            ],
        )
        self.assertEqual(
            [item.bulletText for item in list_items],
            ["•", "•", "•", "°", "°"],
        )
        self.assertEqual(list_items[2].text, "<b>164</b> Resources")
        self.assertEqual(list_items[3].text, "<b>25</b> Analytical Tools")
        self.assertLess(
            list_items[2].style.leftIndent,
            list_items[3].style.leftIndent,
        )

    def test_loose_nested_lists_return_to_top_level_for_each_parent(self):
        release = self.generator.parse_markdown_releases(
            LOOSE_NESTED_LIST_MARKDOWN
        )[0]
        elements = self.generator.parse_html_content(release["fullText"])
        list_items = [
            element
            for element in elements
            if element.style.name in {"ListItem", "NestedListItem"}
        ]

        self.assertEqual(
            [item.style.name for item in list_items],
            ["ListItem", "NestedListItem", "ListItem", "NestedListItem"],
        )
        self.assertEqual(
            [item.text for item in list_items],
            [
                "African Cancer Registry Network",
                "<b>New Dataset</b> - AFCRN Database",
                "Alaska Native Tumor Registry",
                "<b>New Dataset</b> - Cancer in Alaska Native People",
            ],
        )


if __name__ == "__main__":
    unittest.main()
