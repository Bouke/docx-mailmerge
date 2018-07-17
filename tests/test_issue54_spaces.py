import unittest
import os
from mailmerge import MailMerge


class Issue58SpacesTest(unittest.TestCase):

    def setUp(self):
        self.path = os.path.dirname(__file__)

    def test_spaces(self):
        with MailMerge(os.path.join(self.path, 'test_spaces.docx')) as document:
            self.assertEqual(document.get_merge_fields(), {
                'Singleword',
                'Hello world',
                'More than one space'})
