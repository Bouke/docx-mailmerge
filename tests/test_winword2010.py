import unittest
import tempfile
from os import path
from xml.etree import ElementTree

from mailmerge import MailMerge
from tests.utils import EtreeMixin


class MacWord2011Test(EtreeMixin, unittest.TestCase):
    def test(self):
        document = MailMerge(path.join(path.dirname(__file__), 'test_winword2010.docx'))
        self.assertEqual(document.get_merge_fields(),
                         set(['Titel', 'Voornaam', 'Achternaam',
                              'Adresregel_1', 'Postcode', 'Plaats',
                              'Provincie', 'Land_of_regio']))

        document.merge(Voornaam='Bouke', Achternaam='Haarsma',
                       Land_of_regio='The Netherlands', Provincie=None,
                       Postcode='9723 ZA', Plaats='Groningen',
                       Adresregel_1='Helperpark 278d', Titel='dhr.')

        with tempfile.NamedTemporaryFile() as outfile:
            document.write(outfile)

        expected_tree = ElementTree.fromstring(
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p '
            'w:rsidR="00886208" w:rsidRDefault="00886208"><w:r><w:t>dhr.</w:t></w:r><w:r '
            'w:rsidRPr="00886208"><w:rPr><w:lang w:val="nl-NL" /></w:rPr><w:t xml:space="preserve"> '
            '</w:t></w:r><w:r><w:t>Bouke</w:t></w:r><w:r w:rsidRPr="00886208"><w:rPr><w:lang w:val="nl-NL" '
            '/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>Haarsma</w:t></w:r></w:p><w:p '
            'w:rsidR="00886208" w:rsidRDefault="00886208" w:rsidRPr="00886208"><w:pPr><w:rPr><w:lang w:val="nl-NL" '
            '/></w:rPr></w:pPr><w:r><w:t>Helperpark 278d</w:t></w:r></w:p><w:p w:rsidR="00886208" '
            'w:rsidRDefault="00886208"><w:r><w:t>9723 ZA</w:t></w:r><w:r w:rsidRPr="00886208"><w:rPr><w:lang '
            'w:val="nl-NL" /></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>Groningen</w:t></w:r><w:r '
            'w:rsidRPr="00886208"><w:rPr><w:lang w:val="nl-NL" /></w:rPr><w:t xml:space="preserve"> '
            '</w:t></w:r><w:r><w:t /></w:r><w:r w:rsidRPr="00886208"><w:rPr><w:lang w:val="nl-NL" /></w:rPr><w:t '
            'xml:space="preserve"> </w:t></w:r><w:r><w:t>The Netherlands</w:t></w:r></w:p><w:p w:rsidR="00886208" '
            'w:rsidRDefault="00886208" /><w:p w:rsidR="00886208" w:rsidRDefault="00886208"><w:r><w:t>Groningen,'
            '</w:t></w:r></w:p><w:p w:rsidR="00886208" w:rsidRDefault="00886208"><w:bookmarkStart w:id="0" '
            'w:name="_GoBack" /><w:bookmarkEnd w:id="0" /></w:p><w:p w:rsidR="004D7161" '
            'w:rsidRDefault="00886208"><w:r><w:t xml:space="preserve">Dear '
            '</w:t></w:r><w:r><w:t>Bouke</w:t></w:r><w:r><w:t>,</w:t></w:r></w:p><w:p w:rsidR="00886208" '
            'w:rsidRDefault="00886208" /><w:p w:rsidR="00886208" w:rsidRDefault="00886208"><w:r><w:t>I hope this '
            'document from WinWord 2010 finds you well.</w:t></w:r></w:p><w:p w:rsidR="00886208" '
            'w:rsidRDefault="00886208" /><w:p w:rsidR="00886208" w:rsidRDefault="00886208"><w:r><w:t>Kind regards,'
            '</w:t></w:r></w:p><w:p w:rsidR="00886208" w:rsidRDefault="00886208" '
            'w:rsidRPr="00886208"><w:pPr><w:rPr><w:lang w:val="en-US" /></w:rPr></w:pPr><w:proofErr '
            'w:type="spellStart" /><w:r><w:t>docx-mailmerge</w:t></w:r><w:proofErr w:type="spellEnd" /><w:r><w:t>'
            '.</w:t></w:r></w:p><w:sectPr w:rsidR="00886208" w:rsidRPr="00886208"><w:pgSz w:h="16838" w:w="11906" '
            '/><w:pgMar w:bottom="1417" w:footer="708" w:gutter="0" w:header="708" w:left="1417" w:right="1417" '
            'w:top="1417" /><w:cols w:space="708" /><w:docGrid w:linePitch="360" /></w:sectPr></w:body></w:document>')

        self.assert_equal_tree(expected_tree, list(document.parts.values())[0].getroot())
