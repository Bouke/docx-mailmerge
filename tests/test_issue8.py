import unittest
from os import path
from mailmerge import MailMerge


class Issue8Test(unittest.TestCase):
    def test(self):
        document = MailMerge(path.join(path.dirname(__file__), 'test_issue8.docx'))
        self.assertEqual(document.get_merge_fields(), {'testfield'})
