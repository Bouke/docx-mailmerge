import unittest
import tempfile
from os import path
from lxml import etree

from mailmerge import MailMerge
from tests.utils import EtreeMixin, get_document_body_part


class MergeTableRowsMultipartTest(EtreeMixin, unittest.TestCase):
    def setUp(self):
        self.document = MailMerge(path.join(path.dirname(__file__), 'test_merge_table_multipart.docx'))
        self.expected_xml = '<w:document xmlns:ns1="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p w:rsidP="00DA4DE0" w:rsidR="00231DA5" w:rsidRDefault="00DA4DE0"><w:pPr><w:pStyle w:val="Ttulo" /></w:pPr><w:r><w:t>Grades</w:t></w:r></w:p><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:r><w:t>Bouke Haarsma</w:t></w:r><w:r w:rsidR="002C29C5"><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidR="00DA4DE0"><w:t>received the grades</w:t></w:r><w:r w:rsidR="002C29C5"><w:t xml:space="preserve"> for </w:t></w:r><w:r><w:t /></w:r><w:r w:rsidR="00DA4DE0"><w:t xml:space="preserve"> in the table below.</w:t></w:r></w:p><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0" /><w:tbl><w:tblPr><w:tblStyle w:val="Sombreadoclaro-nfasis1" /><w:tblW w:type="auto" w:w="0" /><w:tblLook w:val="04E0" /></w:tblPr><w:tblGrid><w:gridCol w:w="1777" /><w:gridCol w:w="4894" /><w:gridCol w:w="1845" /></w:tblGrid><w:tr w:rsidR="00DA4DE0" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:val="100000000000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0"><w:r><w:t>Class Code</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0"><w:pPr><w:cnfStyle w:val="100000000000" /></w:pPr><w:r><w:t>Class Name</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0"><w:pPr><w:cnfStyle w:val="100000000000" /></w:pPr><w:r><w:t>Grade</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR="00DA4DE0" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:val="000000100000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:r><w:t>ECON101</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:pPr><w:cnfStyle w:val="000000100000" /></w:pPr><w:r><w:t>Economics 101</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:pPr><w:cnfStyle w:val="000000100000" /></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR="00DA4DE0" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:val="000000100000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:r><w:t>ECONADV</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:pPr><w:cnfStyle w:val="000000100000" /></w:pPr><w:r><w:t>Economics Advanced</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:pPr><w:cnfStyle w:val="000000100000" /></w:pPr><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR="00DA4DE0" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:val="000000100000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:r><w:t>OPRES</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:pPr><w:cnfStyle w:val="000000100000" /></w:pPr><w:r><w:t>Operations Research</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="0037127B"><w:pPr><w:cnfStyle w:val="000000100000" /></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR="00C829DD" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:val="010000000000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00C829DD" w:rsidRDefault="00C829DD"><w:r><w:t>THESIS</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00C829DD" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:val="010000000000" /></w:pPr><w:r><w:t>Final thesis</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p w:rsidP="00DA4DE0" w:rsidR="00C829DD" w:rsidRDefault="0037127B"><w:pPr><w:cnfStyle w:val="010000000000" /></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0" w:rsidRPr="00DA4DE0"><w:bookmarkStart w:id="0" w:name="_GoBack" /><w:bookmarkEnd w:id="0" /></w:p><w:sectPr w:rsidR="00DA4DE0" w:rsidRPr="00DA4DE0" w:rsidSect="003B4151"><w:headerReference ns1:id="rId7" w:type="default" /><w:pgSz w:h="16840" w:w="11900" /><w:pgMar w:bottom="1440" w:footer="708" w:gutter="0" w:header="708" w:left="1800" w:right="1800" w:top="1672" /><w:cols w:space="708" /><w:docGrid w:linePitch="360" /></w:sectPr></w:body></w:document>'
        self.expected_tree = etree.fromstring(self.expected_xml) 

    def test_merge_rows_on_multipart_file(self):
        self.assertEqual(self.document.get_merge_fields(),
                         {'student_name', 'study_name', 'class_name', 'class_code', 'class_grade', 'thesis_grade'})

        self.document.merge(
            student_name='Bouke Haarsma',
            study='Industrial Engineering and Management',
            thesis_grade='A',
        )

        self.document.merge_rows('class_code', [
            {'class_code': 'ECON101', 'class_name': 'Economics 101', 'class_grade': 'A'},
            {'class_code': 'ECONADV', 'class_name': 'Economics Advanced', 'class_grade': 'B'},
            {'class_code': 'OPRES', 'class_name': 'Operations Research', 'class_grade': 'A'},
        ])

        with tempfile.TemporaryFile() as outfile:
            self.document.write(outfile)

        self.assert_equal_tree(self.expected_tree, get_document_body_part(self.document).getroot())

    def test_merge_unified_on_multipart_file(self):
        self.document.merge(
            student_name='Bouke Haarsma',
            study='Industrial Engineering and Management',
            thesis_grade='A',
            class_code=[
                {'class_code': 'ECON101', 'class_name': 'Economics 101', 'class_grade': 'A'},
                {'class_code': 'ECONADV', 'class_name': 'Economics Advanced', 'class_grade': 'B'},
                {'class_code': 'OPRES', 'class_name': 'Operations Research', 'class_grade': 'A'},
            ]
        )

        with tempfile.TemporaryFile() as outfile:
            self.document.write(outfile)

        self.assert_equal_tree(self.expected_tree, get_document_body_part(self.document).getroot())

    def tearDown(self):
        self.document.close()
