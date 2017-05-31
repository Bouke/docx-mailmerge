import unittest
import tempfile
from os import path
from lxml import etree

from mailmerge import MailMerge, NAMESPACES
from tests.utils import EtreeMixin


class MergeTableRowsTest(EtreeMixin, unittest.TestCase):
    def setUp(self):
        self.document = MailMerge(path.join(path.dirname(__file__), 'test_merge_table_rows.docx'))
        self.expected_tree = etree.fromstring('<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:ns1="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" mc:Ignorable="w14 wp14"><w:body><w:p ns1:paraId="28ACC80D" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00231DA5" w:rsidRDefault="00DA4DE0"><w:pPr><w:pStyle w:val="Title" /></w:pPr><w:proofErr w:type="spellStart" /><w:r><w:t>Grades</w:t></w:r><w:proofErr w:type="spellEnd" /></w:p><w:p ns1:paraId="07836F52" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:r><w:t>Bouke Haarsma</w:t></w:r><w:r w:rsidR="002C29C5"><w:t xml:space="preserve"> </w:t></w:r><w:proofErr w:type="spellStart" /><w:r w:rsidR="00DA4DE0"><w:t>received</w:t></w:r><w:proofErr w:type="spellEnd" /><w:r w:rsidR="00DA4DE0"><w:t xml:space="preserve"> the </w:t></w:r><w:proofErr w:type="spellStart" /><w:r w:rsidR="00DA4DE0"><w:t>grades</w:t></w:r><w:proofErr w:type="spellEnd" /><w:r w:rsidR="002C29C5"><w:t xml:space="preserve"> </w:t></w:r><w:proofErr w:type="spellStart" /><w:r w:rsidR="002C29C5"><w:t>for</w:t></w:r><w:proofErr w:type="spellEnd" /><w:r w:rsidR="002C29C5"><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t /></w:r><w:r w:rsidR="00DA4DE0"><w:t xml:space="preserve"> in the </w:t></w:r><w:proofErr w:type="spellStart" /><w:r w:rsidR="00DA4DE0"><w:t>table</w:t></w:r><w:proofErr w:type="spellEnd" /><w:r w:rsidR="00DA4DE0"><w:t xml:space="preserve"> below.</w:t></w:r></w:p><w:p ns1:paraId="58525309" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0" /><w:tbl><w:tblPr><w:tblStyle w:val="LightShading-Accent1" /><w:tblW w:type="auto" w:w="0" /><w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="1" w:noHBand="0" w:noVBand="1" w:val="04E0" /></w:tblPr><w:tblGrid><w:gridCol w:w="1777" /><w:gridCol w:w="4894" /><w:gridCol w:w="1845" /></w:tblGrid><w:tr ns1:paraId="599252DD" ns1:textId="77777777" w:rsidR="00DA4DE0" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="1" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="100000000000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="1" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p ns1:paraId="69486D96" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0"><w:r><w:t>Class Code</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p ns1:paraId="1AA11439" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="1" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="100000000000" /></w:pPr><w:r><w:t>Class Name</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p ns1:paraId="3091846F" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="1" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="100000000000" /></w:pPr><w:r><w:t>Grade</w:t></w:r></w:p></w:tc></w:tr><w:tr ns1:paraId="1699D5B6" ns1:textId="77777777" w:rsidR="00DA4DE0" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="1" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p ns1:paraId="5D81DF7F" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:r><w:t>ECON101</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p ns1:paraId="5A67E49A" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:pPr><w:r><w:t>Economics 101</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p ns1:paraId="5EB9BD23" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr ns1:paraId="1699D5B6" ns1:textId="77777777" w:rsidR="00DA4DE0" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="1" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p ns1:paraId="5D81DF7F" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:r><w:t>ECONADV</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p ns1:paraId="5A67E49A" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:pPr><w:r><w:t>Economics Advanced</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p ns1:paraId="5EB9BD23" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:pPr><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr><w:tr ns1:paraId="1699D5B6" ns1:textId="77777777" w:rsidR="00DA4DE0" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="1" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p ns1:paraId="5D81DF7F" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:r><w:t>OPRES</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p ns1:paraId="5A67E49A" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:pPr><w:r><w:t>Operations Research</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p ns1:paraId="5EB9BD23" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="1" w:oddVBand="0" w:val="000000100000" /></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr><w:tr ns1:paraId="0B5730FD" ns1:textId="77777777" w:rsidR="00C829DD" w:rsidTr="00C829DD"><w:trPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="1" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="010000000000" /></w:trPr><w:tc><w:tcPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="1" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="001000000000" /><w:tcW w:type="dxa" w:w="1809" /></w:tcPr><w:p ns1:paraId="6A211A5A" ns1:textId="4E90EB38" w:rsidP="00DA4DE0" w:rsidR="00C829DD" w:rsidRDefault="00C829DD"><w:r><w:t>THESIS</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="5529" /></w:tcPr><w:p ns1:paraId="12FCE443" ns1:textId="289CA6C6" w:rsidP="00DA4DE0" w:rsidR="00C829DD" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="1" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="010000000000" /></w:pPr><w:proofErr w:type="spellStart" /><w:r><w:t>Final</w:t></w:r><w:proofErr w:type="spellEnd" /><w:r><w:t xml:space="preserve"> thesis</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1178" /></w:tcPr><w:p ns1:paraId="0ACD9198" ns1:textId="413A3A4D" w:rsidP="00DA4DE0" w:rsidR="00C829DD" w:rsidRDefault="00C829DD"><w:pPr><w:cnfStyle w:evenHBand="0" w:evenVBand="0" w:firstColumn="0" w:firstRow="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastColumn="0" w:lastRow="1" w:lastRowFirstColumn="0" w:lastRowLastColumn="0" w:oddHBand="0" w:oddVBand="0" w:val="010000000000" /></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p ns1:paraId="26A87379" ns1:textId="77777777" w:rsidP="00DA4DE0" w:rsidR="00DA4DE0" w:rsidRDefault="00DA4DE0" w:rsidRPr="00DA4DE0"><w:bookmarkStart w:id="0" w:name="_GoBack" /><w:bookmarkEnd w:id="0" /></w:p><w:sectPr w:rsidR="00DA4DE0" w:rsidRPr="00DA4DE0" w:rsidSect="002C0659"><w:pgSz w:h="16840" w:w="11900" /><w:pgMar w:bottom="1440" w:footer="708" w:gutter="0" w:header="708" w:left="1800" w:right="1800" w:top="1440" /><w:cols w:space="708" /><w:docGrid w:linePitch="360" /></w:sectPr></w:body></w:document>')  # noqa

    def test_merge_rows(self):
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

        self.assert_equal_tree(self.expected_tree,
                               list(self.document.parts.values())[0].getroot())

    def test_merge_rows_no_table(self):
        """
        Merging of rows should not fail if there is no table with the given
        column in the document
        """
        self.document.merge(
            student_name='Bouke Haarsma',
            study='Industrial Engineering and Management',
            thesis_grade='A',
            class_code=[
                {'class_code': 'ECON101', 'class_name': 'Economics 101', 'class_grade': 'A'},
                {'class_code': 'ECONADV', 'class_name': 'Economics Advanced', 'class_grade': 'B'},
                {'class_code': 'OPRES', 'class_name': 'Operations Research', 'class_grade': 'A'},
            ],
            no_table=[
                {'no_table': 'Table not available'}
            ]
        )

        with tempfile.TemporaryFile() as outfile:
            self.document.write(outfile)

        self.assert_equal_tree(self.expected_tree,
                               list(self.document.parts.values())[0].getroot())

    def test_merge_rows_remove_table(self):
        """
        When flag is set and there is no data in a table, table needs to be removed
        """
        self.document.remove_empty_tables = True
        self.document.merge(
            student_name='Bouke Haarsma',
            study='Industrial Engineering and Management',
            thesis_grade='A',
            class_code=[]
        )

        with tempfile.TemporaryFile() as outfile:
            self.document.write(outfile)
        self.assertIsNone(
            list(self.document.parts.values())[0].getroot().find('.//{%(w)s}tbl' % NAMESPACES)
        )

    def test_merge_unified(self):
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

        self.assert_equal_tree(self.expected_tree,
                               list(self.document.parts.values())[0].getroot())

    def tearDown(self):
        self.document.close()
