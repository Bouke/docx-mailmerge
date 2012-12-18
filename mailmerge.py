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

        r = re.compile(r' MERGEFIELD (.+?) \\\* MERGEFORMAT ', re.I)
        for part in self.parts.values():
            # Remove attribute that soft-links to other namespaces; other namespaces
            # are not used, so would cause word to throw an error.
            del part.getroot().attrib['{%(mc)s}Ignorable' % NAMESPACES]

            for parent in part.iterfind('.//{%(w)s}fldSimple/..' % NAMESPACES):
                for idx, child in enumerate(parent):
                    if child.tag != '{%(w)s}fldSimple' % NAMESPACES:
                        continue
                    instr = child.attrib['{%(w)s}instr' % NAMESPACES]
                    m = r.match(instr)
                    parent[idx] = Element('MergeField', name=m.group(1))

            for parent in part.iterfind('.//{%(w)s}instrText/../..' % NAMESPACES):
                for idx, child in enumerate(parent):
                    fldChar = child.find('{%(w)s}fldChar[@{%(w)s}fldCharType="begin"]' % NAMESPACES)
                    if fldChar is None:
                        continue
                    instr = parent[idx + 1].find('{%(w)s}instrText' % NAMESPACES)
                    if instr is None:
                        continue
                    m = r.match(instr.text)
                    if m is None:
                        continue
                    parent[idx] = Element('MergeField', name=m.group(1))
                    to_delete += [(parent, parent[i])
                                  for i in range(idx+1, idx+5)]

        for parent, child in to_delete:
            parent.remove(child)

    def write(self, file):
        for field in self.get_merge_fields():
            self.merge(**{field:''})

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

    def merge(self, parts=None, **replacements):
        if not parts:
            parts = self.parts.values()

        for part in parts:
            for field, text in replacements.items():
                self.__merge_field(part, field, text)

    def __merge_field(self, part, field, text):
        for mf in part.findall('.//MergeField[@name="%s"]' % field):
            mf.clear()
            mf.tag = '{%(w)s}r' % NAMESPACES
            mf.append(ElementTree.Element('{%(w)s}t' % NAMESPACES))
            mf[0].text = text

    def find_row_node(self, field, parts=None):
        if not parts:
            parts = self.parts.values()
        for part in parts:
            for table in part.iterfind('.//{%(w)s}tbl' % NAMESPACES):
                for idx, row in enumerate(table):
                    if row.find('.//MergeField[@name="%s"]' % field) is not None:
                        return table, idx, row
