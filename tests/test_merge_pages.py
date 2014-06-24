import unittest
import tempfile
from os import path
from xml.etree import ElementTree

from mailmerge import MailMerge
from tests.utils import EtreeMixin, get_document_body_part


class MergeListTest(EtreeMixin, unittest.TestCase):
    def test_pages(self):
        document = MailMerge(path.join(path.dirname(__file__), 'test_merge_pages.docx'))
        self.assertEqual(document.get_merge_fields(), {'fieldname'})

        document.merge_pages([
            {'fieldname': "xyz"},
            {'fieldname': "abc"},
            {'fieldname': "2b v ~2b"},
       ])

        with tempfile.TemporaryFile() as outfile:
            document.write(outfile)

        expected_tree = ElementTree.fromstring(
            '<w:document xmlns:ns1="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p w:rsidP="007B2B98"'
            ' w:rsidR="00EC3BBB" w:rsidRDefault="00651722" w:rsidRPr="00651722"><w:r><w:t xml:space="preserve">This'
            ' is a template for the </w:t></w:r><w:r><w:t>xyz</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list'
            ' test case</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack" /><w:bookmarkEnd w:id="0" /></w:p><w:sectPr'
            ' w:rsidR="00EC3BBB" w:rsidRPr="00651722"><w:headerReference ns1:id="rId7" w:type="default" /><w:footerReference'
            ' ns1:id="rId8" w:type="default" /><w:pgSz w:h="15840" w:w="12240" /><w:pgMar w:bottom="1440" w:footer="708"'
            ' w:gutter="0" w:header="708" w:left="1440" w:right="1440" w:top="1440" /><w:cols w:space="708" /><w:docGrid'
            ' w:linePitch="360" /></w:sectPr></w:body><w:br w:type="page" /><w:body><w:p w:rsidP="007B2B98"'
            ' w:rsidR="00EC3BBB" w:rsidRDefault="00651722" w:rsidRPr="00651722"><w:r><w:t xml:space="preserve">This'
            ' is a template for the </w:t></w:r><w:r><w:t>abc</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list'
            ' test case</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack" /><w:bookmarkEnd w:id="0" /></w:p><w:sectPr'
            ' w:rsidR="00EC3BBB" w:rsidRPr="00651722"><w:headerReference ns1:id="rId7" w:type="default"'
            ' /><w:footerReference ns1:id="rId8" w:type="default" /><w:pgSz w:h="15840" w:w="12240" /><w:pgMar'
            ' w:bottom="1440" w:footer="708" w:gutter="0" w:header="708" w:left="1440" w:right="1440" w:top="1440"'
            ' /><w:cols w:space="708" /><w:docGrid w:linePitch="360" /></w:sectPr></w:body><w:br w:type="page"'
            ' /><w:body><w:p w:rsidP="007B2B98" w:rsidR="00EC3BBB" w:rsidRDefault="00651722" w:rsidRPr="00651722">'
            '<w:r><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>2b v ~2b</w:t></w:r><w:r>'
            '<w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"'
            ' /><w:bookmarkEnd w:id="0" /></w:p><w:sectPr w:rsidR="00EC3BBB" w:rsidRPr="00651722"><w:headerReference'
            ' ns1:id="rId7" w:type="default" /><w:footerReference ns1:id="rId8" w:type="default" /><w:pgSz w:h="15840"'
            ' w:w="12240" /><w:pgMar w:bottom="1440" w:footer="708" w:gutter="0" w:header="708" w:left="1440"'
            ' w:right="1440" w:top="1440" /><w:cols w:space="708" /><w:docGrid w:linePitch="360" /></w:sectPr>'
            '</w:body></w:document>')

        self.assert_equal_tree(expected_tree, get_document_body_part(document).getroot())


    def test_pages_with_multiple_pages(self):
        """
        This also tests that merge_pages produces a multiple paged document,
        however this template already contains two pages. So the result should
        be 3 * 2 pages.
        """
        document = MailMerge(path.join(path.dirname(__file__), 'test_merge_pages_paged.docx'))
        self.assertEqual(document.get_merge_fields(), {'fieldname'})

        document.merge_pages([
            {'fieldname': "xyz"},
            {'fieldname': "abc"},
            {'fieldname': "2b v ~2b"},
       ])

        with tempfile.TemporaryFile() as outfile:
            document.write(outfile)

        expected_tree = ElementTree.fromstring(
            '<w:document xmlns:ns1="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p w:rsidP="007B2B98"'
            ' w:rsidR="00507D2F" w:rsidRDefault="00651722"><w:r><w:t xml:space="preserve">This is a template for'
            ' the </w:t></w:r><w:r><w:t>xyz</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t>'
            '</w:r><w:r w:rsidR="00507D2F"><w:t>, page 1</w:t></w:r></w:p><w:p w:rsidR="00507D2F" w:rsidRDefault="00507D2F">'
            '<w:r><w:br w:type="page" /></w:r></w:p><w:p w:rsidP="007B2B98" w:rsidR="00EC3BBB" w:rsidRDefault="00507D2F"'
            ' w:rsidRPr="00651722"><w:r><w:lastRenderedPageBreak /><w:t xml:space="preserve">This is a template for the'
            ' </w:t></w:r><w:r><w:t>xyz</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:r>'
            '<w:t>, page 2</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack" /><w:bookmarkEnd w:id="0" /></w:p>'
            '<w:sectPr w:rsidR="00EC3BBB" w:rsidRPr="00651722"><w:headerReference ns1:id="rId7" w:type="default" />'
            '<w:footerReference ns1:id="rId8" w:type="default" /><w:pgSz w:h="15840" w:w="12240" /><w:pgMar'
            ' w:bottom="1440" w:footer="708" w:gutter="0" w:header="708" w:left="1440" w:right="1440" w:top="1440"'
            ' /><w:cols w:space="708" /><w:docGrid w:linePitch="360" /></w:sectPr></w:body><w:br w:type="page" />'
            '<w:body><w:p w:rsidP="007B2B98" w:rsidR="00507D2F" w:rsidRDefault="00651722"><w:r><w:t xml:space="preserve"'
            '>This is a template for the </w:t></w:r><w:r><w:t>abc</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list'
            ' test case</w:t></w:r><w:r w:rsidR="00507D2F"><w:t>, page 1</w:t></w:r></w:p><w:p w:rsidR="00507D2F"'
            ' w:rsidRDefault="00507D2F"><w:r><w:br w:type="page" /></w:r></w:p><w:p w:rsidP="007B2B98" w:rsidR="00EC3BBB"'
            ' w:rsidRDefault="00507D2F" w:rsidRPr="00651722"><w:r><w:lastRenderedPageBreak /><w:t xml:space="preserve">'
            'This is a template for the </w:t></w:r><w:r><w:t>abc</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list'
            ' test case</w:t></w:r><w:r><w:t>, page 2</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack" />'
            '<w:bookmarkEnd w:id="0" /></w:p><w:sectPr w:rsidR="00EC3BBB" w:rsidRPr="00651722"><w:headerReference'
            ' ns1:id="rId7" w:type="default" /><w:footerReference ns1:id="rId8" w:type="default" /><w:pgSz w:h="15840"'
            ' w:w="12240" /><w:pgMar w:bottom="1440" w:footer="708" w:gutter="0" w:header="708" w:left="1440"'
            ' w:right="1440" w:top="1440" /><w:cols w:space="708" /><w:docGrid w:linePitch="360" /></w:sectPr></w:body>'
            '<w:br w:type="page" /><w:body><w:p w:rsidP="007B2B98" w:rsidR="00507D2F" w:rsidRDefault="00651722"><w:r>'
            '<w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>2b v ~2b</w:t></w:r><w:r><w:t'
            ' xml:space="preserve"> merge_list test case</w:t></w:r><w:r w:rsidR="00507D2F"><w:t>, page 1</w:t></w:r>'
            '</w:p><w:p w:rsidR="00507D2F" w:rsidRDefault="00507D2F"><w:r><w:br w:type="page" /></w:r></w:p><w:p'
            ' w:rsidP="007B2B98" w:rsidR="00EC3BBB" w:rsidRDefault="00507D2F" w:rsidRPr="00651722"><w:r>'
            '<w:lastRenderedPageBreak /><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>2b'
            ' v ~2b</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:r><w:t>, page 2</w:t>'
            '</w:r><w:bookmarkStart w:id="0" w:name="_GoBack" /><w:bookmarkEnd w:id="0" /></w:p><w:sectPr w:rsidR="00EC3BBB"'
            ' w:rsidRPr="00651722"><w:headerReference ns1:id="rId7" w:type="default" /><w:footerReference ns1:id="rId8"'
            ' w:type="default" /><w:pgSz w:h="15840" w:w="12240" /><w:pgMar w:bottom="1440" w:footer="708" w:gutter="0"'
            ' w:header="708" w:left="1440" w:right="1440" w:top="1440" /><w:cols w:space="708" /><w:docGrid'
            ' w:linePitch="360" /></w:sectPr></w:body></w:document>')

        self.assert_equal_tree(expected_tree, get_document_body_part(document).getroot())

