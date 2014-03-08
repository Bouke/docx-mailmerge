import unittest
from mailmerge import MailMerge


class Issue2Test(unittest.TestCase):
    def test(self):
        MailMerge("tests/test_issue2.docx")
