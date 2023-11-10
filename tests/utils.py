import zipfile
import io
import sys
import tempfile
from os import path

from lxml import etree
from mailmerge import MailMerge, NAMESPACES, CONTENT_TYPES_PARTS

TEXTS_XPATH = "//w:t/text()"

class EtreeMixin(object):

    IGNORED_FIELDS=[
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR"
        ,"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr"
        ]

    def filter_item(self, item):
        return item[0] not in self.IGNORED_FIELDS

    def assert_equal_tree(self, lhs, rhs):
        """
        Compares two instances of ElementTree are equivalent.
        """
        self.assertEqual(lhs.tag, rhs.tag)
        self.assertEqual(len(lhs), len(rhs))
        self.assertEqual(lhs.text or '', rhs.text or '')
        self.assertEqual(sorted(filter(self.filter_item, lhs.items())), sorted(filter(self.filter_item, rhs.items())))
        for lhs_child, rhs_child in zip(lhs, rhs):
            self.assert_equal_tree(lhs_child, rhs_child)

    def assert_equal_tree_debug(self, lhs, rhs):
        try:
            self.assert_equal_tree(lhs, rhs)
        except:
            with open("expected.xml", "wb") as f:
                f.write(etree.tostring(lhs.getroottree().getroot(), encoding='UTF-8', pretty_print=True))
            with open("real.xml", "wb") as f:
                f.write(etree.tostring(rhs.getroottree().getroot(), encoding='UTF-8', pretty_print=True))
            raise
    
    def merge_templates(self, filename, replacements, separator='page_break', write_file=True, mm_args=[], mm_kwargs={}, mt_args=[], mt_kwargs={}, output=None):
        with MailMerge(path.join(path.dirname(__file__), filename), *mm_args, **mm_kwargs) as document:
            document.merge_templates(replacements, separator, *mt_args, **mt_kwargs)

            if write_file:
                if output:
                    with open(output, 'wb') as outfile:
                        document.write(outfile)
                else:
                    with tempfile.TemporaryFile() as outfile:
                        document.write(outfile)

            return document, get_document_body_part(document).getroot()

    def merge(self, filename, replacements, write_file=True, mm_args=[], mm_kwargs={}, output=None):
        with MailMerge(path.join(path.dirname(__file__), filename), *mm_args, **mm_kwargs) as document:
            document.merge(**replacements)

            if write_file:
                if output:
                    with open(output, 'wb') as outfile:
                        document.write(outfile)
                else:
                    with tempfile.TemporaryFile() as outfile:
                        document.write(outfile)

            return document, get_document_body_part(document).getroot()

    def open_docx(self, filename):
        self.docx_zipfile = zipfile.ZipFile(filename)
        self.docx_parts = {}

        content_types = etree.parse(self.docx_zipfile.open('[Content_Types].xml'))
        for file_elem in content_types.findall('{%(ct)s}Override' % NAMESPACES):
            part_type = file_elem.attrib['ContentType' % NAMESPACES]
            if part_type in CONTENT_TYPES_PARTS:
                fn = file_elem.attrib['PartName' % NAMESPACES].split('/', 1)[1]
                zi = self.docx_zipfile.getinfo(fn)
                self.docx_parts[zi] = etree.parse(self.docx_zipfile.open(zi))

    def get_part(self, endswidth='}document'):
        for zi, part in self.docx_parts.items():
            if part.getroot().tag.endswith('}document'):
                return zi, part
        return None, None

    def get_new_docx(self, replacement_parts):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as output:
            for zi in self.docx_zipfile.filelist:
                if zi in replacement_parts:
                    xml = etree.tostring(replacement_parts[zi].getroot(), encoding='UTF-8', xml_declaration=True)
                    output.writestr(zi.filename, xml)
                else:
                    output.writestr(zi.filename, self.docx_zipfile.read(zi))
        return zip_buffer


def get_document_body_part(document, endswith="document"):
    for part in document.parts.values():
        if part['part'].getroot().tag.endswith('}%s' % endswith):
            return part['part']

    raise AssertionError("main document body not found in document.parts")

def get_document_body_parts(document, endswith="document"):
    parts = []
    for part in document.parts.values():
        if part['part'].getroot().tag.endswith('}%s' % endswith):
            parts.append(part['part'])
    for _, _, part in document.new_parts:
        if part.getroot().tag.endswith('}%s' % endswith):
            parts.append(part)

    return parts
