import unittest
import tempfile
import warnings
import datetime
import locale
from os import path
from lxml import etree

from mailmerge import NAMESPACES
from tests.utils import EtreeMixin, get_document_body_part

MERGE_FIELDS_TRUE_XPATH = './w:mailMerge'

class MergeParamsTest(EtreeMixin, unittest.TestCase):

    def test_merge_params_all(self):
        """ tests if all fields are merged """
        values = {"first": "one"}
        document, root_elem = self.merge(
            'test_merge_params.docx',
            values,
            mm_kwargs={"merge_params": "all"},
            output="tests/output/test_output_merge_params_all.docx"
            )

        # print(document._has_unmerged_fields, document.get_merge_fields())
        self.assertListEqual(
            document.settings.getroot().xpath(MERGE_FIELDS_TRUE_XPATH, namespaces=NAMESPACES),
            [])

    def test_merge_params_some(self):
        """ tests if all fields are merged """
        values = {"first": "one"}
        document, root_elem = self.merge(
            'test_merge_params.docx',
            values,
            mm_kwargs={"merge_params": "some"},
            output="tests/output/test_output_merge_params_some.docx"
            )

        self.assertListEqual(
            document.settings.getroot().xpath(MERGE_FIELDS_TRUE_XPATH, namespaces=NAMESPACES),
            [])
        # print(document._has_unmerged_fields, document.get_merge_fields())

    def test_merge_params_none(self):
        """ tests if all fields are merged """
        values = {"first": "one"}
        document, root_elem = self.merge(
            'test_merge_params.docx',
            values,
            mm_kwargs={"merge_params": "none"},
            output="tests/output/test_output_merge_params_none.docx"
            )

        self.assertListEqual(
            document.settings.getroot().xpath(MERGE_FIELDS_TRUE_XPATH, namespaces=NAMESPACES),
            [])
        # print(document._has_unmerged_fields, document.get_merge_fields())
