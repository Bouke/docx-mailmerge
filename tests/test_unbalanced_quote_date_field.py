import unittest
from os import path

from mailmerge import MailMerge


class EmptyFieldTest(unittest.TestCase):
    def test(self):
        with MailMerge(path.join(path.dirname(__file__), "test_unbalanced_quote_date_field.docx")) as document:
            self.assertEqual(document.get_merge_fields(), set())
