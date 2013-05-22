import unittest
from mailmerge import MailMerge


class Issue2Test(unittest.TestCase):
    def test(self):
        self.assertRaisesRegexp(
            ValueError, 'Could not determine name of merge '
                        'field in value " MERGEFIELD  Title "',
            MailMerge, 'tests/test_issue2.docx')
