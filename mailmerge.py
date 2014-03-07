from copy import deepcopy
import re

# Why lxml? XPath! Plus the more rational and simple xmlns preservation
# Not to mention that lxml and xml are mostly compatible.
# Oh, and faster
from lxml import etree
from lxml.etree import ElementTree
from lxml.etree import Element
from zipfile import ZipFile, ZIP_DEFLATED

INTERNAL_NAMESPACE = "internal:docx-mailmerge"

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'int': INTERNAL_NAMESPACE,
    'xml': "http://www.w3.org/XML/1998/namespace",
}

CONTENT_TYPES = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
)

def first(lst):
    for elem in lst:
        return elem
    return None

class MailMerge(object):
    def __init__(self, file):
        self.zip = ZipFile(file)
        self.parts = {}

        content_types = ElementTree(file=self.zip.open('[Content_Types].xml'))
        for file in content_types.iterfind('ct:Override', namespaces=NAMESPACES):
            type = file.attrib['ContentType']
            if type in CONTENT_TYPES:
                fn = file.attrib['PartName'].split('/', 1)[1]
                zi = self.zip.getinfo(fn)
                self.parts[zi] = ElementTree(file=self.zip.open(zi))

        r = re.compile(r'\s*MERGEFIELD\s+"?([^\s"]+?)"?\s+(|\\\*\s+MERGEFORMAT)\s*', re.I)
        for part in self.parts.values():
            # Remove attribute that soft-links to other namespaces; other namespaces
            # are not used, so would cause word to throw an error.
            ignorable_key = '{%(mc)s}Ignorable' % NAMESPACES
            if ignorable_key in part.getroot().attrib:
                part.getroot().attrib.pop(ignorable_key)

            for parent in part.iterfind('.//w:fldSimple/..', namespaces=NAMESPACES):

                for idx, child in enumerate(parent.iterfind("w:fldSimple", namespaces=NAMESPACES)):
                    instr = child.xpath('@w:instr', namespaces=NAMESPACES)[0]

                    m = r.match(instr)
                    if not m:
                        raise ValueError('Could not determine name of merge '
                                         'field in value "%s"' % instr)

                    # Extract original w:r structure to preserve formatting
                    childspan = child.xpath('w:r', namespaces=NAMESPACES)[0]
                    childtext = first(childspan.xpath('w:t', namespaces=NAMESPACES))
                    if childtext is None:
                        childtext = Element("{%(w)s}t" % NAMESPACES)
                        childspan.append(childtext)
                    childtext.set("{%(xml)s}space" % NAMESPACES, "preserve")
                    childtext.set("{%(int)s}merge-field-name" % NAMESPACES, m.group(1))
                    parent.insert(parent.index(child),childspan)
                    parent.remove(child)
                    
            # Eliminate duplicate iteration with set

            for parent in set(part.iterfind('.//w:instrText/../..', namespaces=NAMESPACES)):
                # Eliminated indices to allow on the fly old element deletion
                fields = zip(
                    parent.findall('w:r/w:fldChar[@w:fldCharType="begin"]/..', namespaces=NAMESPACES),
                    parent.findall('w:r/w:fldChar[@w:fldCharType="end"]/..', namespaces=NAMESPACES),
                    parent.findall('w:r/w:instrText/..', namespaces=NAMESPACES))
                for beginElm, endElm, instrElm in fields:
                    instrFld = instrElm.find("w:instrText", namespaces=NAMESPACES)
                    idx_begin = parent.index(beginElm)
                    idx_end = parent.index(endElm)
                    idx_instr = parent.index(instrElm)
                    instr = instrFld.text
                    m = r.match(instr)
                    if m is None:
                        continue

                    # Original code caused following error due to premature erasing of the initial element, and repeated call to the same parent as a result of xpath evaluation for parent
                    # New code solves the problem by deleting the fields early
                    # and by eliminating duplicate iterations
                    if not (idx_begin < idx_instr and idx_instr < idx_end):
                        raise ValueError("Invalid word document containing stray instrText element without containing fldChar elements: %d, %d, %d: %s" % (idx_begin, idx_instr, idx_end, instr))

                    # Preserve formatting
                    # Preserve instrElm
                    # But remove instrFld
                    instrElm.remove(instrFld)
                    # Append new w:t if not present
                    textFld = first(instrElm.xpath("w:t", namespaces=NAMESPACES))
                    if textFld is None:
                        textFld = Element("{%(w)s}t" % NAMESPACES)
                        textFld.set("{%(xml)s}space" % NAMESPACES, "preserve")
                        instrElm.append(textFld)
                    
                    textFld.set("{%(int)s}merge-field-name" % NAMESPACES, m.group(1))
                    # Add deletion first
                    
                    deleteSet = [parent[i] for i in range(idx_begin, idx_instr)] + [parent[i] for i in range(idx_instr + 1, idx_end + 1)]
                    
                    for elem in deleteSet:
                        parent.remove(elem)

    def write(self, file):
        for field in self.get_merge_fields():
            self.merge(**{field:''})
        
        output = ZipFile(file, 'w', ZIP_DEFLATED)
        for zi in self.zip.filelist:
            if zi in self.parts:
                xml = etree.tostring(self.parts[zi].getroot())
                output.writestr(zi.filename, xml)
            else:
                output.writestr(zi.filename, self.zip.read(zi))
        output.close()

    def get_merge_fields(self, parts=None):
        if not parts:
            parts = self.parts.values()
        fields = set()
        for part in parts:
            for mf in part.findall('.//w:t[@int:merge-field-name]', namespaces=NAMESPACES):
                fields.add(mf.attrib['{%(int)s}merge-field-name' % NAMESPACES])
        return fields

    def merge(self, parts=None, **replacements):
        if not parts:
            parts = self.parts.values()

        for part in parts:
            for field, text in replacements.items():
                self.__merge_field(part, field, text)

    def __merge_field(self, part, field, text):
        for mf in part.findall('.//w:t[@int:merge-field-name="%s"]' % (field, ), namespaces=NAMESPACES):
            mf.attrib.pop("{%(int)s}merge-field-name" % NAMESPACES)
            
            oldmf = mf
            mf = Element(oldmf.tag)
            for key, val in oldmf.attrib.iteritems():
                mf.set(key, val)
            mf.text = text
            
            # Replacing the element is the only way of clearing nsmap
            # See https://bugs.launchpad.net/lxml/+bug/555602
            container = oldmf.getparent()
            container.insert(container.index(oldmf), mf)
            container.remove(oldmf)

    def merge_rows(self, anchor, rows):
        table, idx, template = self.__find_row_anchor(anchor)
        del table[idx]
        for row_data in rows:
            row = deepcopy(template)
            self.merge([row], **row_data)
            table.append(row)

    def __find_row_anchor(self, field, parts=None):
        if not parts:
            parts = self.parts.values()
        for part in parts:
            for table in part.iterfind('.//w:tbl', namespaces=NAMESPACES):
                for idx, row in enumerate(table):
                    if row.find('.//w:t[@int:merge-field-name="%s"]' % (field, ), namespaces=NAMESPACES) is not None:
                        return table, idx, row
