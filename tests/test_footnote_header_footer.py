import unittest
import tempfile
import warnings
import datetime
import locale
from os import path
from lxml import etree

from mailmerge import NAMESPACES
from tests.utils import EtreeMixin, get_document_body_part

FOOTNOTE_XPATH = "//w:footnote[@w:id = '1']/w:p/w:r/w:t/text()"

class FootnoteHeaderFooterTest(EtreeMixin, unittest.TestCase):

    @unittest.expectedFailure
    def test_all(self):
        values = ["one", "two", "three"]
        # header/footer/footnotes don't work with multiple replacements, only with merge
        # fix this when it is implemented
        document, root_elem = self.merge_templates(
            'test_footnote_header_footer.docx',
            [{"fieldname": value, "footerfield": "f_" + value, "headerfield": "h_" + value}
                for value in values
                ],
            # output="tests/test_output_footnote_header_footer.docx"
            )

        footnote_root_elem = get_document_body_part(document, "footnotes").getroot()
        footnote = "".join(footnote_root_elem.xpath(FOOTNOTE_XPATH, namespaces=NAMESPACES))
        correct_footnote = " Merge : one "
        self.assertEqual(footnote, correct_footnote)

    def test_only_merge(self):
        values = ["one", "two", "three"]
        document, root_elem = self.merge(
            'test_footnote_header_footer.docx',
            [{"fieldname": value, "footerfield": "f_" + value, "headerfield": "h_" + value}
                for value in values
                ][0],
            # output="tests/test_output_one_footnote_header_footer.docx"
            )

        footnote_root_elem = get_document_body_part(document, "footnotes").getroot()
        footnote = "".join(footnote_root_elem.xpath(FOOTNOTE_XPATH, namespaces=NAMESPACES))
        self.assertEqual(footnote, " Merge : one")
