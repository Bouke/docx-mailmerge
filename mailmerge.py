from copy import deepcopy
import re
from xml.etree import ElementTree
from xml.etree.ElementTree import Element
from zipfile import ZipFile, ZIP_DEFLATED

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}

for prefix, uri in NAMESPACES.items():
    ElementTree.register_namespace(prefix, uri)

CONTENT_TYPES = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
)


class MailMerge(object):
    def __init__(self, file):
        self.zip = ZipFile(file)
        self.parts = {}

        content_types = ElementTree.parse(self.zip.open('[Content_Types].xml'))
        for file in content_types.iterfind('{%(ct)s}Override' % NAMESPACES):
            type = file.attrib['ContentType' % NAMESPACES]
            if type in CONTENT_TYPES:
                fn = file.attrib['PartName' % NAMESPACES].split('/', 1)[1]
                zi = self.zip.getinfo(fn)
                self.parts[zi] = ElementTree.parse(self.zip.open(zi))

        to_delete = []

        r = re.compile(r' MERGEFIELD +"?([^ ]+?)"? +(|\\\* MERGEFORMAT )', re.I)
        for part in self.parts.values():
            # Remove attribute that soft-links to other namespaces; other namespaces
            # are not used, so would cause word to throw an error.
            ignorable_key = '{%(mc)s}Ignorable' % NAMESPACES
            if ignorable_key in part.getroot().attrib:
                del part.getroot().attrib[ignorable_key]

            for parent in part.iterfind('.//{%(w)s}fldSimple/..' % NAMESPACES):
                for idx, child in enumerate(parent):
                    if child.tag != '{%(w)s}fldSimple' % NAMESPACES:
                        continue
                    instr = child.attrib['{%(w)s}instr' % NAMESPACES]

                    m = r.match(instr)
                    if not m:
                        raise ValueError('Could not determine name of merge '
                                         'field in value "%s"' % instr)
                    parent[idx] = Element('MergeField', name=m.group(1))

            for parent in part.iterfind('.//{%(w)s}instrText/../..' % NAMESPACES):
                children = list(parent)
                fields = zip(
                    [children.index(e) for e in
                     parent.findall('{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="begin"]/..' % NAMESPACES)],
                    [children.index(e) for e in
                     parent.findall('{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="end"]/..' % NAMESPACES)],
                    [e.text for e in
                     parent.findall('{%(w)s}r/{%(w)s}instrText' % NAMESPACES)])
                for idx_begin, idx_end, instr in fields:
                    m = r.match(instr)
                    if m is None:
                        continue
                    parent[idx_begin] = Element('MergeField', name=m.group(1))
                    to_delete += [(parent, parent[i + 1])
                                  for i in range(idx_begin, idx_end)]

        for parent, child in to_delete:
            parent.remove(child)

    def write(self, file):
        # Replace all remaining merge fields with empty values
        for field in self.get_merge_fields():
            self.merge(**{field: ''})

        output = ZipFile(file, 'w', ZIP_DEFLATED)
        for zi in self.zip.filelist:
            if zi in self.parts:
                xml = ElementTree.tostring(self.parts[zi].getroot())
                output.writestr(zi.filename, xml)
            else:
                output.writestr(zi.filename, self.zip.read(zi))
        output.close()

    def get_merge_fields(self, parts=None):
        if not parts:
            parts = self.parts.values()
        fields = set()
        for part in parts:
            for mf in part.findall('.//MergeField'):
                fields.add(mf.attrib['name'])
        return fields

    def merge_pages(self, replacements):
        """
        Duplicate template page. Creates a copy of the template for each item
        in the list, does a merge, and separates the them by page breaks.
        """
        for part in self.parts.values():
            root = part.getroot()

            tag = root.tag
            if tag == '{%(w)s}ftr' % NAMESPACES or tag == '{%(w)s}hdr' % NAMESPACES:
                continue

            children = []
            for child in root:
                root.remove(child)
                children.append(child)

            for i, repl in enumerate(replacements):
                # Add page break in between replacements
                if i > 0:
                    pagebreak = ElementTree.Element('{%(w)s}br' % NAMESPACES)
                    pagebreak.attrib['{%(w)s}type' % NAMESPACES] = 'page'
                    root.append(pagebreak)

                parts = []
                for child in children:
                    child_copy = deepcopy(child)
                    root.append(child_copy)
                    parts.append(child_copy)
                self.merge(parts, **repl)

    def merge(self, parts=None, **replacements):
        if not parts:
            parts = self.parts.values()

        for part in parts:
            for field, replacement in replacements.items():
                if isinstance(replacement, list):
                    self.merge_rows(field, replacement)
                else:
                    self.__merge_field(part, field, replacement)

    def __merge_field(self, part, field, text):
        for mf in part.findall('.//MergeField[@name="%s"]' % field):
            mf.clear()
            mf.tag = '{%(w)s}r' % NAMESPACES
            mf.append(ElementTree.Element('{%(w)s}t' % NAMESPACES))
            mf[0].text = text

    def merge_rows(self, anchor, rows):
        table, idx, template = self.__find_row_anchor(anchor)
        del table[idx]
        for i, row_data in enumerate(rows):
            row = deepcopy(template)
            self.merge([row], **row_data)
            table.insert(idx + i, row)

    def __find_row_anchor(self, field, parts=None):
        if not parts:
            parts = self.parts.values()
        for part in parts:
            for table in part.iterfind('.//{%(w)s}tbl' % NAMESPACES):
                for idx, row in enumerate(table):
                    if row.find('.//MergeField[@name="%s"]' % field) is not None:
                        return table, idx, row
