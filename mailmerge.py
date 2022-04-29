from copy import deepcopy
import shlex
from xml.etree import ElementTree as ET
import warnings
from zipfile import ZipFile, ZIP_DEFLATED

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

CONTENT_TYPES_PARTS = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
)

CONTENT_TYPE_SETTINGS = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
)


class MailMerge(object):
    def __init__(self, file, remove_empty_tables=False):
        self.zip = ZipFile(file)
        self.parts = {}
        self.settings = None
        self._settings_info = None
        self.remove_empty_tables = remove_empty_tables

        try:
            content_types = ET.parse(self.zip.open("[Content_Types].xml"))
            for file in content_types.findall("{%(ct)s}Override" % NAMESPACES):
                type = file.attrib["ContentType" % NAMESPACES]
                if type in CONTENT_TYPES_PARTS:
                    zi, self.parts[zi] = self.__get_tree_of_file(file)
                elif type == CONTENT_TYPE_SETTINGS:
                    self._settings_info, self.settings = self.__get_tree_of_file(file)

            to_delete = []

            for part in self.parts.values():

                for parent in part.findall(".//{%(w)s}fldSimple/.." % NAMESPACES):
                    for idx, child in enumerate(parent):
                        if child.tag != "{%(w)s}fldSimple" % NAMESPACES:
                            continue
                        instr = child.attrib["{%(w)s}instr" % NAMESPACES]

                        name = self.__parse_instr(instr)
                        if name is None:
                            continue
                        parent[idx] = ET.Element("MergeField", name=name)

                for parent in part.findall(".//{%(w)s}instrText/../.." % NAMESPACES):
                    children = list(parent)
                    fields = zip(
                        [
                            children.index(e)
                            for e in parent.findall(
                                '{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="begin"]/..'
                                % NAMESPACES
                            )
                        ],
                        [
                            children.index(e)
                            for e in parent.findall(
                                '{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="end"]/..'
                                % NAMESPACES
                            )
                        ],
                    )

                    for idx_begin, idx_end in fields:
                        # consolidate all instrText nodes between 'begin' and 'end' into a single node
                        begin = children[idx_begin]
                        instrText_elements = [
                            e
                            for e in parent.findall(
                                "{%(w)s}r/{%(w)s}instrText" % NAMESPACES
                            )
                        ]
                        instrText_parent_elements = [
                            e
                            for e in parent.findall(
                                "{%(w)s}r/{%(w)s}instrText/.." % NAMESPACES
                            )
                        ]
                        instr_elements = [
                            e
                            for e in zip(instrText_elements, instrText_parent_elements)
                            if idx_begin < children.index(e[1]) < idx_end
                        ]

                        # instr_elements = [e for e in parent.findall('{%(w)s}r/{%(w)s}instrText' % NAMESPACES)
                        #                  if idx_begin < children.index(parent) < idx_end]
                        if len(instr_elements) == 0:
                            continue

                        # set the text of the first instrText element to the concatenation
                        # of all the instrText element texts
                        instr_text = "".join([e[0].text for e in instr_elements])
                        instr_elements[0][0].text = instr_text

                        # delete all instrText elements except the first
                        for instr in instr_elements[1:]:
                            instr[1].remove(instr[0])

                        name = self.__parse_instr(instr_text)
                        if name is None:
                            continue

                        parent[idx_begin] = ET.Element("MergeField", name=name)

                        # use this so we know *where* to put the replacement
                        instr_elements[0][0].tag = "MergeText"
                        block = instr_elements[0][1]
                        # append the other tags in the w:r block too
                        parent[idx_begin].extend(list(block))

                        to_delete += [
                            (parent, parent[i + 1]) for i in range(idx_begin, idx_end)
                        ]

            for parent, child in to_delete:
                parent.remove(child)

            # Remove mail merge settings to avoid error messages when opening document in Winword
            if self.settings:
                settings_root = self.settings.getroot()
                mail_merge = settings_root.find("{%(w)s}mailMerge" % NAMESPACES)
                if mail_merge is not None:
                    settings_root.remove(mail_merge)
        except:
            self.zip.close()
            raise

    @classmethod
    def __parse_instr(cls, instr):
        args = shlex.split(instr, posix=False)
        if args[0] != "MERGEFIELD":
            return None
        name = args[1]
        if name[0] == '"' and name[-1] == '"':
            name = name[1:-1]
        return name

    def __get_tree_of_file(self, file):
        fn = file.attrib["PartName" % NAMESPACES].split("/", 1)[1]
        zi = self.zip.getinfo(fn)
        return zi, ET.parse(self.zip.open(zi))

    def write(self, file):
        # Replace all remaining merge fields with empty values
        for field in self.get_merge_fields():
            self.merge(**{field: ""})

        with ZipFile(file, "w", ZIP_DEFLATED) as output:
            for zi in self.zip.filelist:
                if zi in self.parts:
                    xml = ET.tostring(self.parts[zi].getroot())
                    output.writestr(zi.filename, xml)
                elif zi == self._settings_info:
                    xml = ET.tostring(self.settings.getroot())
                    output.writestr(zi.filename, xml)
                else:
                    output.writestr(zi.filename, self.zip.read(zi))

    def get_merge_fields(self, parts=None):
        if not parts:
            parts = self.parts.values()
        fields = set()
        for part in parts:
            for mf in part.findall(".//MergeField"):
                fields.add(mf.attrib["name"])
        return fields

    def merge_templates(self, replacements, separator):
        """
        Duplicate template. Creates a copy of the template, does a merge, and separates them by a new paragraph, a new break or a new section break.
        separator must be :
        - page_break : Page Break.
        - column_break : Column Break. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - textWrapping_break : Line Break.
        - continuous_section : Continuous section break. Begins the section on the next paragraph.
        - evenPage_section : evenPage section break. section begins on the next even-numbered page, leaving the next odd page blank if necessary.
        - nextColumn_section : nextColumn section break. section begins on the following column on the page. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - nextPage_section : nextPage section break. section begins on the following page.
        - oddPage_section : oddPage section break. section begins on the next odd-numbered page, leaving the next even page blank if necessary.
        """

        # TYPE PARAM CONTROL AND SPLIT
        valid_separators = {
            "page_break",
            "column_break",
            "textWrapping_break",
            "continuous_section",
            "evenPage_section",
            "nextColumn_section",
            "nextPage_section",
            "oddPage_section",
        }
        if not separator in valid_separators:
            raise ValueError("Invalid separator argument")
        type, sepClass = separator.split("_")

        # GET ROOT - WORK WITH DOCUMENT
        for part in self.parts.values():
            root = part.getroot()
            tag = root.tag
            if tag == "{%(w)s}ftr" % NAMESPACES or tag == "{%(w)s}hdr" % NAMESPACES:
                continue

            if sepClass == "section":

                # FINDING FIRST SECTION OF THE DOCUMENT
                firstSection_element = root.find(
                    "w:body/w:p/w:pPr/w:sectPr", namespaces=NAMESPACES
                )
                firstSection_parent = root.find(
                    "w:body/w:p/w:pPr/w:sectPr/..", namespaces=NAMESPACES
                )
                firstSection = (firstSection_element, firstSection_parent)

                if firstSection[0] is None:
                    firstSection_element = root.find(
                        "w:body/w:sectPr", namespaces=NAMESPACES
                    )
                    firstSection_parent = root.find(
                        "w:body/w:sectPr/..", namespaces=NAMESPACES
                    )
                    firstSection = (firstSection_element, firstSection_parent)

                # MODIFY TYPE ATTRIBUTE OF FIRST SECTION FOR MERGING
                # if firstSection[0] is not None:
                nextPageSec = deepcopy(firstSection[0])
                for child in nextPageSec:
                    # Delete old type if exist
                    if child.tag == "{%(w)s}type" % NAMESPACES:
                        nextPageSec.remove(child)
                # Create new type (def parameter)
                newType = ET.SubElement(nextPageSec, "{%(w)s}type" % NAMESPACES)
                newType.set("{%(w)s}val" % NAMESPACES, type)

                # REPLACING FIRST SECTION
                sec_index = list(firstSection[1]).index(firstSection[0])
                secRoot = firstSection[1]
                secRoot.remove(firstSection[0])
                secRoot.insert(sec_index, nextPageSec)

            # FINDING LAST SECTION OF THE DOCUMENT
            lastSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)

            # SAVING LAST SECTION
            mainSection = deepcopy(lastSection)
            lsecRoot = root.find("w:body/w:sectPr/..", namespaces=NAMESPACES)
            lsecRoot.remove(lastSection)

            # COPY CHILDREN ELEMENTS OF BODY IN A LIST
            childrenList = root.findall("w:body/*", namespaces=NAMESPACES)

            # DELETE ALL CHILDREN OF BODY
            for child in root:
                if child.tag == "{%(w)s}body" % NAMESPACES:
                    child.clear()

            # REFILL BODY AND MERGE DOCS - ADD LAST SECTION ENCAPSULATED OR NOT
            lr = len(replacements)
            lc = len(childrenList)

            for i, repl in enumerate(replacements):
                parts = []
                for (j, n) in enumerate(childrenList):
                    element = deepcopy(n)
                    for child in root:
                        if child.tag == "{%(w)s}body" % NAMESPACES:
                            child.append(element)
                            parts.append(element)
                            if (j + 1) == lc:
                                if (i + 1) == lr:
                                    child.append(mainSection)
                                    parts.append(mainSection)
                                else:
                                    if sepClass == "section":
                                        intSection = deepcopy(mainSection)
                                        p = ET.SubElement(
                                            child, "{%(w)s}p" % NAMESPACES
                                        )
                                        pPr = ET.SubElement(
                                            p, "{%(w)s}pPr" % NAMESPACES
                                        )
                                        pPr.append(intSection)
                                        parts.append(p)
                                    elif sepClass == "break":
                                        pb = ET.SubElement(
                                            child, "{%(w)s}p" % NAMESPACES
                                        )
                                        r = ET.SubElement(pb, "{%(w)s}r" % NAMESPACES)
                                        nbreak = ET.Element("{%(w)s}br" % NAMESPACES)
                                        nbreak.attrib["{%(w)s}type" % NAMESPACES] = type
                                        r.append(nbreak)

                    self.merge(parts, **repl)

    def merge_pages(self, replacements):
        """
        Deprecated method.
        """
        warnings.warn(
            "merge_pages has been deprecated in favour of merge_templates",
            category=DeprecationWarning,
            stacklevel=2,
        )
        self.merge_templates(replacements, "page_break")

    def merge(self, parts=None, **replacements):
        if not parts:
            parts = self.parts.values()

        for field, replacement in replacements.items():
            if isinstance(replacement, list):
                self.merge_rows(field, replacement)
            else:
                for part in parts:
                    self.__merge_field(part, field, replacement)

    def __merge_field(self, part, field, text):
        for mf in part.findall('.//MergeField[@name="%s"]' % field):
            children = list(mf)
            mf.clear()  # clear away the attributes
            mf.tag = "{%(w)s}r" % NAMESPACES
            mf.extend(children)

            nodes = []
            # preserve new lines in replacement text
            text = text or ""  # text might be None
            text_parts = str(text).replace("\r", "").split("\n")
            for i, text_part in enumerate(text_parts):
                text_node = ET.Element("{%(w)s}t" % NAMESPACES)
                text_node.text = text_part
                nodes.append(text_node)

                # if not last node add new line node
                if i < (len(text_parts) - 1):
                    nodes.append(ET.Element("{%(w)s}br" % NAMESPACES))

            ph = mf.find("MergeText")
            if ph is not None:
                # add text nodes at the exact position where
                # MergeText was found
                index = list(mf).index(ph)
                for node in reversed(nodes):
                    mf.insert(index, node)
                mf.remove(ph)
            else:
                mf.extend(nodes)

    def merge_rows(self, anchor, rows):
        table, idx, template, table_parent = self.__find_row_anchor(anchor)
        if table is not None:
            if len(rows) > 0:
                del table[idx]
                for i, row_data in enumerate(rows):
                    row = deepcopy(template)
                    self.merge([row], **row_data)
                    table.insert(idx + i, row)
            else:
                # if there is no data for a given table
                # we check whether table needs to be removed
                if self.remove_empty_tables:
                    table_parent.remove(table)

    def __find_row_anchor(self, field, parts=None):
        if not parts:
            parts = self.parts.values()
        for part in parts:
            table_elements = [x for x in part.findall(".//{%(w)s}tbl" % NAMESPACES)]
            table_parents = [x for x in part.findall(".//{%(w)s}tbl/.." % NAMESPACES)]

            tables = zip(table_elements, table_parents)
            for table in tables:
                for idx, row in enumerate(table[0]):
                    if row.find('.//MergeField[@name="%s"]' % field) is not None:
                        return table[0], idx, row, table[1]
        return None, None, None, None

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def close(self):
        if self.zip is not None:
            try:
                self.zip.close()
            finally:
                self.zip = None
