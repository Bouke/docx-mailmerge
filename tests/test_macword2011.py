import unittest
import tempfile
from mailmerge import MailMerge


class MacWord2011Test(unittest.TestCase):
    def test(self):
        document = MailMerge('tests/test_macword2011.docx')
        self.assertEquals(document.get_merge_fields(),
                          set(['first_name', 'last_name', 'country', 'state',
                               'postal_code', 'date', 'address_line', 'city']))

        document.merge(first_name='Bouke', last_name='Haarsma',
                       country='The Netherlands', state=None,
                       postal_code='9723 ZA', city='Groningen',
                       address_line='Helperpark 278d', date='May 22nd, 2013')

        with tempfile.NamedTemporaryFile() as outfile:
            document.write(outfile)
