import unittest

from mailmerge import NAMESPACES
from tests.utils import EtreeMixin, get_document_body_part, TEXTS_XPATH, get_document_body_parts

FOOTNOTE_XPATH = "//w:footnote[@w:id = '1']/w:p/w:r/w:t/text()"

class FootnoteHeaderFooterTest(EtreeMixin, unittest.TestCase):
    # @TODO test missing values
    # @TODO test if separator isn't section
    # @TODO test headers/footers with relations
    @unittest.expectedFailure
    def test_all(self):
        values = ["one", "two", "three"]
        # header/footer/footnotes don't work with multiple replacements, only with merge
        # fix this when it is implemented
        document, root_elem = self.merge_templates(
            'test_footnote_header_footer.docx',
            [
                {
                "fieldname": value,
                "footerfield": "f_" + value,
                "headerfield": "h_" + value,
                "footerfirst": "ff_" + value,
                "headerfirst": "hf_" + value,
                "footereven": "fe_" + value,
                "headereven": "he_" + value,
                }
                for value in values
                ],
            separator="nextPage_section",
            # output="tests/output/test_output_footnote_header_footer.docx"
            )

        footers = sorted([
            "".join(footer_doc_tree.getroot().xpath(TEXTS_XPATH, namespaces=NAMESPACES))
            for footer_doc_tree in get_document_body_parts(document, endswith="ftr")
        ])
        self.assertListEqual(footers, sorted([
            footer
            for value in values
            for footer in [
            'Footer on even page fe_%s' % value, 
            'Footer on every page f_%s' % value,
            'Footer on first page ff_%s' % value]] + [
            'Footer on even page ', 
            'Footer on every page ',
            'Footer on first page ']
            ))

        headers = sorted([
            "".join(header_doc_tree.getroot().xpath(TEXTS_XPATH, namespaces=NAMESPACES))
            for header_doc_tree in get_document_body_parts(document, endswith="hdr")
        ])
        self.assertListEqual(headers, sorted([
            header
            for value in values
            for header in [
            'Header even: he_%s' % value,
            'Header on every page: h_%s' % value,
            'Header on first page: hf_%s' % value]] + [
            'Header even: ', 
            'Header on every page: ',
            'Header on first page: ']
            ))

        footnote_root_elem = get_document_body_part(document, "footnotes").getroot()
        footnote = "".join(footnote_root_elem.xpath(FOOTNOTE_XPATH, namespaces=NAMESPACES))
        correct_footnote = " Merge : one "
        self.assertEqual(footnote, correct_footnote)

    def test_only_merge(self):
        values = ["one", "two", "three"]
        document, root_elem = self.merge(
            'test_footnote_header_footer.docx',
            [
                {
                "fieldname": value,
                "footerfield": "f_" + value,
                "headerfield": "h_" + value,
                "footerfirst": "ff_" + value,
                "headerfirst": "hf_" + value,
                "footereven": "fe_" + value,
                "headereven": "he_" + value,
                }
                for value in values
                ][0],
            # output="tests/output/test_output_one_footnote_header_footer.docx"
            )

        footnote_root_elem = get_document_body_part(document, "footnotes").getroot()
        footnote = "".join(footnote_root_elem.xpath(FOOTNOTE_XPATH, namespaces=NAMESPACES))
        self.assertEqual(footnote, " Merge : one")

        footers = sorted([
            "".join(footer_doc_tree.getroot().xpath(TEXTS_XPATH, namespaces=NAMESPACES))
            for footer_doc_tree in get_document_body_parts(document, endswith="ftr")
        ])
        value = values[0]
        self.assertListEqual(footers, [
            'Footer on even page fe_%s' % value, 
            'Footer on every page f_%s' % value,
            'Footer on first page ff_%s' % value])

        headers = sorted([
            "".join(header_doc_tree.getroot().xpath(TEXTS_XPATH, namespaces=NAMESPACES))
            for header_doc_tree in get_document_body_parts(document, endswith="hdr")
        ])
        self.assertListEqual(headers, [
            'Header even: he_%s' % value,
            'Header on every page: h_%s' % value,
            'Header on first page: hf_%s' % value])


    def test_footer(self):
        values = ["one", "two"]
        # header/footer/footnotes don't work with multiple replacements, only with merge
        # fix this when it is implemented
        document, root_elem = self.merge_templates(
            'test_footer.docx',
            [
                {
                "footer": value,
                }
                for value in values
                ],
            separator="nextPage_section",
            # output="tests/output/test_footer.docx"
            )
