import unittest
import tempfile
from os import path
from lxml import etree

from mailmerge import MailMerge
from tests.utils import EtreeMixin, get_document_body_part


class MultipleElementsTest(EtreeMixin, unittest.TestCase):
    def test(self):
        with MailMerge(path.join(path.dirname(__file__), 'test_multiple_elements.docx')) as document:
            self.assertEqual(document.get_merge_fields(),
                             set(['foo', 'bar', 'gak']))

            document.merge(foo='one', bar='two', gak='three')

            with tempfile.TemporaryFile() as outfile:
                document.write(outfile)

            expected_tree = etree.fromstring(
                '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
                'xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" '
                'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                'xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
                'xmlns:v="urn:schemas-microsoft-com:vml" '
                'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
                'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                'xmlns:w10="urn:schemas-microsoft-com:office:word" '
                'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
                'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
                'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
                'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
                'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
                'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 '
                'wp14"><w:body><w:p w14:paraId="0670CD65" w14:textId="7CABA628" w:rsidR="00575A14" '
                'w:rsidRDefault="004923E9" w:rsidP="00EA5502"><w:pPr><w:rPr><w:sz '
                'w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>one</w:t></w:r></w:p><w:p '
                'w14:paraId="3DDDF43F" w14:textId="60BBA4C0" w:rsidR="009B730B" w:rsidRDefault="002A768F" '
                'w:rsidP="00EA5502"><w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz '
                'w:val="22"/></w:rPr><w:t>two</w:t></w:r></w:p><w:p w14:paraId="3CADC6B4" w14:textId="174E4BC7" '
                'w:rsidR="004B5BC1" w:rsidRPr="00335531" w:rsidRDefault="00CC2A10" '
                'w:rsidP="00EA5502"><w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz '
                'w:val="22"/></w:rPr><w:t>three</w:t></w:r></w:p><w:sectPr w:rsidR="004B5BC1" w:rsidRPr="00335531" '
                'w:rsidSect="0090608D"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
                'w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols '
                'w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')

            self.assert_equal_tree(expected_tree, get_document_body_part(document).getroot())
