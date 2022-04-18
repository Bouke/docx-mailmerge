import unittest
import tempfile
from os import path
from lxml import etree

from mailmerge import MailMerge, NAMESPACES
from tests.utils import EtreeMixin, get_document_body_part

class NextRecordsTest(EtreeMixin, unittest.TestCase):
    """
    Testing next records
    """
    
    def test_next_record(self):
        """
        Tests if the next record field works
        """
        values = ["one", "two", "three", "four", "five"]
        document, root_elem = self.merge_templates(
            'test_next_record.docx',
            [{"field": value}
                for value in values
                ],
            # output="tests/test_output_next_record.docx"
            )
        
        self.assertFalse(root_elem.xpath("//MergeField", namespaces=NAMESPACES))
        fields = root_elem.xpath("//w:t/text()", namespaces=NAMESPACES)
        expected = [
            v
            for value in values + ['']*3
            for v in [value, '/', value]
        ]
        self.assertListEqual(fields, expected)
