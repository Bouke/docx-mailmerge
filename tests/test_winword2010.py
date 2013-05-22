import unittest
import tempfile
from mailmerge import MailMerge


class MacWord2011Test(unittest.TestCase):
    def test(self):
        document = MailMerge('tests/test_winword2010.docx')
        self.assertEquals(document.get_merge_fields(),
                          set(['Titel', 'Voornaam', 'Achternaam',
                               'Adresregel_1', 'Postcode', 'Plaats',
                               'Provincie', 'Land_of_regio']))

        document.merge(Voornaam='Bouke', Achternaam='Haarsma',
                       Land_of_regio='The Netherlands', Provincie=None,
                       Postcode='9723 ZA', Plaats='Groningen',
                       Adresregel_1='Helperpark 278d', Titel='dhr.')

        with tempfile.NamedTemporaryFile() as outfile:
            document.write(outfile.name)
