import unittest
from os import path
from mailmerge import MailMerge


class Issue64Test(unittest.TestCase):
    def test(self):
        with MailMerge(path.join(path.dirname(__file__), 'test_issue8.docx')) as document:
            document.merge(testfield=10)
