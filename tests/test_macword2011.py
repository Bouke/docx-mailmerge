import unittest
import tempfile
from os import path
from lxml import etree

from mailmerge import MailMerge
from tests.utils import EtreeMixin, get_document_body_part


class MacWord2011Test(EtreeMixin, unittest.TestCase):
    def test(self):
        with MailMerge(path.join(path.dirname(__file__), 'test_macword2011.docx')) as document:
            self.assertEqual(document.get_merge_fields(),
                             set(['first_name', 'last_name', 'country', 'state',
                                  'postal_code', 'date', 'address_line', 'city']))

            document.merge(first_name='Bouke', last_name='Haarsma',
                           country='The Netherlands', state=None,
                           postal_code='9723 ZA', city='Groningen',
                           address_line='Helperpark 278d', date='May 22nd, 2013')

            with tempfile.TemporaryFile() as outfile:
                document.write(outfile)

        expected_tree = etree.fromstring(
            '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:ns1="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w="http://schemas'
            '.openxmlformats.org/wordprocessingml/2006/main" mc:Ignorable="w14 wp14">'
            '<w:body><w:p ns1:paraId="25CDEC86" ns1:textId="77777777" '
            'w:rsidR="00FB567A" w:rsidRDefault="00541684"><w:r><w:t>Bouke</w:t></w:r><w:r w:rsidR="00916690"><w:t '
            'xml:space="preserve"> </w:t></w:r><w:r><w:t>Haarsma</w:t></w:r></w:p><w:p ns1:paraId="67F7A559" '
            'ns1:textId="77777777" w:rsidR="00916690" w:rsidRDefault="00541684"><w:r><w:t>Helperpark '
            '278d</w:t></w:r></w:p><w:p ns1:paraId="228F2AEA" ns1:textId="77777777" w:rsidR="00916690" '
            'w:rsidRDefault="00541684"><w:r><w:t>9723 ZA</w:t></w:r><w:r><w:t xml:space="preserve"> '
            '</w:t></w:r><w:r><w:t>Groningen</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack" /><w:bookmarkEnd '
            'w:id="0" /><w:r w:rsidR="00916690"><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t /></w:r><w:r '
            'w:rsidR="00916690"><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>The '
            'Netherlands</w:t></w:r></w:p><w:p ns1:paraId="696E15D7" ns1:textId="77777777" w:rsidR="00916690" '
            'w:rsidRDefault="00916690" /><w:p ns1:paraId="3F48DA35" ns1:textId="77777777" w:rsidR="00916690" '
            'w:rsidRDefault="00916690" /><w:p ns1:paraId="68E7FBD1" ns1:textId="77777777" w:rsidR="00916690" '
            'w:rsidRDefault="00916690"><w:r><w:t xml:space="preserve">Groningen, </w:t></w:r><w:r><w:t>May 22nd, '
            '2013</w:t></w:r><w:r><w:t>,</w:t></w:r></w:p><w:p ns1:paraId="26B47597" ns1:textId="77777777" '
            'w:rsidR="00916690" w:rsidRDefault="00916690" /><w:p ns1:paraId="740FF493" ns1:textId="77777777" '
            'w:rsidR="00916690" w:rsidRDefault="00916690" /><w:p ns1:paraId="232699A0" ns1:textId="77777777" '
            'w:rsidR="00916690" w:rsidRDefault="00916690"><w:r><w:t xml:space="preserve">Dear '
            '</w:t></w:r><w:r><w:t>Bouke</w:t></w:r><w:r><w:t>,</w:t></w:r></w:p><w:p ns1:paraId="1AE04A19" '
            'ns1:textId="77777777" w:rsidR="00916690" w:rsidRDefault="00916690" /><w:p ns1:paraId="298752B0" '
            'ns1:textId="77777777" w:rsidR="00916690" w:rsidRDefault="00916690" /><w:p ns1:paraId="477BE6DB" '
            'ns1:textId="77777777" w:rsidR="00916690" w:rsidRDefault="00916690"><w:r><w:t>I hope this message finds '
            'you well.</w:t></w:r></w:p><w:p ns1:paraId="77542CE2" ns1:textId="77777777" w:rsidR="00916690" '
            'w:rsidRDefault="00916690" /><w:p ns1:paraId="734AC552" ns1:textId="77777777" w:rsidR="00916690" '
            'w:rsidRDefault="00916690" /><w:p ns1:paraId="1680D84C" ns1:textId="77777777" w:rsidR="00916690" '
            'w:rsidRDefault="00916690"><w:r><w:t>Kind regards,</w:t></w:r></w:p><w:p ns1:paraId="67F698AC" '
            'ns1:textId="77777777" w:rsidR="00916690" w:rsidRDefault="00916690" /><w:p ns1:paraId="7A4C001B" '
            'ns1:textId="77777777" w:rsidR="00916690" w:rsidRDefault="00916690"><w:r><w:t>docx-mailmerge'
            '.</w:t></w:r></w:p><w:sectPr w:rsidR="00916690" w:rsidSect="00CA5A66"><w:pgSz w:h="16840" w:w="11900" '
            '/><w:pgMar w:bottom="1417" w:footer="708" w:gutter="0" w:header="708" w:left="1417" w:right="1417" '
            'w:top="1417" /><w:cols w:space="708" /><w:docGrid w:linePitch="360" /></w:sectPr></w:body></w:document>')

        self.assert_equal_tree(expected_tree, get_document_body_part(document).getroot())
