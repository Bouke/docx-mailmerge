import unittest
import tempfile
import warnings
from os import path
from lxml import etree

from mailmerge import MailMerge
from tests.utils import EtreeMixin, get_document_body_part


class MergeListTest(EtreeMixin, unittest.TestCase):
    def test_pages(self):
        with MailMerge(path.join(path.dirname(__file__), 'test_merge_pages.docx')) as document:
            self.assertEqual(document.get_merge_fields(), {'fieldname'})

            with warnings.catch_warnings(record=True) as warning_list:
                document.merge_pages([
                    {'fieldname': "xyz"},
                    {'fieldname': "abc"},
                    {'fieldname': "2b v ~2b"},
                ])
                self.assertTrue(any(item.category == DeprecationWarning for item in warning_list))

            with tempfile.TemporaryFile() as outfile:
                document.write(outfile)

        expected_tree = etree.fromstring('<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body><w:p w:rsidR="00EC3BBB" w:rsidRPr="00651722" w:rsidRDefault="00651722" w:rsidP="007B2B98"><w:r><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>xyz</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p w:rsidR="00EC3BBB" w:rsidRPr="00651722" w:rsidRDefault="00651722" w:rsidP="007B2B98"><w:r><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>abc</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p w:rsidR="00EC3BBB" w:rsidRPr="00651722" w:rsidRDefault="00651722" w:rsidP="007B2B98"><w:r><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>2b v ~2b</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:sectPr w:rsidR="00EC3BBB" w:rsidRPr="00651722"><w:headerReference w:type="default" r:id="rId7"/><w:footerReference w:type="default" r:id="rId8"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')  # noqa

        self.assert_equal_tree(expected_tree, get_document_body_part(document).getroot())

    def test_pages_with_multiple_pages(self):
        """
        This also tests that merge_pages produces a multiple paged document,
        however this template already contains two pages. So the result should
        be 3 * 2 pages.
        """
        with MailMerge(path.join(path.dirname(__file__), 'test_merge_pages_paged.docx')) as document:
            self.assertEqual(document.get_merge_fields(), {'fieldname'})

            with warnings.catch_warnings(record=True) as warning_list:
                document.merge_pages([
                    {'fieldname': "xyz"},
                    {'fieldname': "abc"},
                    {'fieldname': "2b v ~2b"},
                ])
                self.assertTrue(any(item.category == DeprecationWarning for item in warning_list))

            with tempfile.TemporaryFile() as outfile:
                document.write(outfile)

        expected_tree = etree.fromstring('<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body><w:p w:rsidR="00507D2F" w:rsidRDefault="00651722" w:rsidP="007B2B98"><w:r><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>xyz</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:r w:rsidR="00507D2F"><w:t>, page 1</w:t></w:r></w:p><w:p w:rsidR="00507D2F" w:rsidRDefault="00507D2F"><w:r><w:br w:type="page"/></w:r></w:p><w:p w:rsidR="00EC3BBB" w:rsidRPr="00651722" w:rsidRDefault="00507D2F" w:rsidP="007B2B98"><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>xyz</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:r><w:t>, page 2</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p w:rsidR="00507D2F" w:rsidRDefault="00651722" w:rsidP="007B2B98"><w:r><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>abc</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:r w:rsidR="00507D2F"><w:t>, page 1</w:t></w:r></w:p><w:p w:rsidR="00507D2F" w:rsidRDefault="00507D2F"><w:r><w:br w:type="page"/></w:r></w:p><w:p w:rsidR="00EC3BBB" w:rsidRPr="00651722" w:rsidRDefault="00507D2F" w:rsidP="007B2B98"><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>abc</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:r><w:t>, page 2</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p w:rsidR="00507D2F" w:rsidRDefault="00651722" w:rsidP="007B2B98"><w:r><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>2b v ~2b</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:r w:rsidR="00507D2F"><w:t>, page 1</w:t></w:r></w:p><w:p w:rsidR="00507D2F" w:rsidRDefault="00507D2F"><w:r><w:br w:type="page"/></w:r></w:p><w:p w:rsidR="00EC3BBB" w:rsidRPr="00651722" w:rsidRDefault="00507D2F" w:rsidP="007B2B98"><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">This is a template for the </w:t></w:r><w:r><w:t>2b v ~2b</w:t></w:r><w:r><w:t xml:space="preserve"> merge_list test case</w:t></w:r><w:r><w:t>, page 2</w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:sectPr w:rsidR="00EC3BBB" w:rsidRPr="00651722"><w:headerReference w:type="default" r:id="rId7"/><w:footerReference w:type="default" r:id="rId8"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')  # noqa

        self.assert_equal_tree(expected_tree, get_document_body_part(document).getroot())
