import unittest
import tempfile
from os import path
from lxml import etree

from mailmerge import MailMerge, NAMESPACES
from tests.utils import EtreeMixin, get_document_body_part

DOCPR_TAG = 'wp:docPr'

class MergeTemplatesWithPicturesTest(EtreeMixin, unittest.TestCase):

    def test_duplicate_docPr(self):
        document, root_elem = self.merge_templates(
            'test_input_duplicate_id.docx',
            [{'field': 'test1'}, {'field': 'test2'}])

        docPr_ids = root_elem.xpath("//wp:docPr/@id", namespaces=NAMESPACES)

        self.assertEqual(len(docPr_ids), len(set(docPr_ids)))
