import unittest

from mailmerge import NAMESPACES
from tests.utils import EtreeMixin, TEXTS_XPATH

MERGE_FIELDS_TRUE_XPATH = './w:mailMerge'
SEPARATE_TEXT_FIELDS_XPATH = '//w:fldChar[@w:fldCharType = "separate"]/../following-sibling::w:r/w:t/text()'
SIMPLE_FIELDS_TEXT_FIELDS_XPATH = '//w:fldSimple/w:r/w:t/text()'
VALUES = {"first": "one", "three_simple": "three_simple"}
TEST_DOCX = 'test_keep_fields.docx'
TEST_DOCX_OUT = "tests/output/test_keep_fields_%s.docx"

class MergeParamsTest(EtreeMixin, unittest.TestCase):
    """ Tests the three values of the keep_fields parameter

    The test is done using 4 MERGEFIELD fields
    one - complex field, having data
    two - complex field, no data
    three_simple - simple field, having data
    four_simple - simple field, no data
    """

    def test_keep_fields_all(self):
        """ tests if all fields are merged """
        keep_fields = "all"
        document, root_elem = self.merge(
            TEST_DOCX,
            VALUES,
            mm_kwargs={"keep_fields": keep_fields},
            # output=TEST_DOCX_OUT % keep_fields
            )

        self.assertListEqual(
            document.get_settings().getroot().xpath(MERGE_FIELDS_TRUE_XPATH, namespaces=NAMESPACES),
            [])
        self.assertListEqual(
            root_elem.xpath(TEXTS_XPATH, namespaces=NAMESPACES),
            ["one", "«second»", 'three_simple', '«four_simple»'])
        self.assertListEqual(
            root_elem.xpath(SEPARATE_TEXT_FIELDS_XPATH, namespaces=NAMESPACES),
            ["one", "«second»"])
        self.assertListEqual(
            root_elem.xpath(SIMPLE_FIELDS_TEXT_FIELDS_XPATH, namespaces=NAMESPACES),
            ["three_simple", "«four_simple»"])

    def test_keep_fields_some(self):
        """ tests if all fields are merged """
        keep_fields = "some"
        document, root_elem = self.merge(
            TEST_DOCX,
            VALUES,
            mm_kwargs={"keep_fields": keep_fields},
            # output=TEST_DOCX_OUT % keep_fields
            )

        self.assertListEqual(
            document.get_settings().getroot().xpath(MERGE_FIELDS_TRUE_XPATH, namespaces=NAMESPACES),
            [])
        self.assertListEqual(
            root_elem.xpath(TEXTS_XPATH, namespaces=NAMESPACES),
            ["one", "«second»", 'three_simple', '«four_simple»'])
        self.assertListEqual(
            root_elem.xpath(SEPARATE_TEXT_FIELDS_XPATH, namespaces=NAMESPACES),
            ["«second»"])
        self.assertListEqual(
            root_elem.xpath(SIMPLE_FIELDS_TEXT_FIELDS_XPATH, namespaces=NAMESPACES),
            ["«four_simple»"])

    def test_keep_fields_none(self):
        """ tests if all fields are merged """
        keep_fields = "none"
        document, root_elem = self.merge(
            TEST_DOCX,
            VALUES,
            mm_kwargs={"keep_fields": keep_fields},
            # output=TEST_DOCX_OUT % keep_fields
            )

        self.assertListEqual(
            document.get_settings().getroot().xpath(MERGE_FIELDS_TRUE_XPATH, namespaces=NAMESPACES),
            [])
        self.assertListEqual(
            root_elem.xpath(TEXTS_XPATH, namespaces=NAMESPACES),
            ["one", "", "three_simple", ""])
        self.assertListEqual(
            root_elem.xpath(SEPARATE_TEXT_FIELDS_XPATH, namespaces=NAMESPACES),
            [])
        self.assertListEqual(
            root_elem.xpath(SIMPLE_FIELDS_TEXT_FIELDS_XPATH, namespaces=NAMESPACES),
            [])
