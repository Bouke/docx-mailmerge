import warnings
import shlex
import re
import datetime
# import locale
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy

from lxml.etree import Element
from lxml import etree


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'xml': 'http://www.w3.org/XML/1998/namespace'
}

CONTENT_TYPES_PARTS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
)

CONTENT_TYPE_SETTINGS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'

VALID_SEPARATORS = {
    'page_break', 'column_break', 'textWrapping_break',
    'continuous_section', 'evenPage_section', 'nextColumn_section', 'nextPage_section', 'oddPage_section'}

NUMBERFORMAT_RE = r"([^0.,'#PN]+)?(P\d+|N\d+|[0.,'#]+%?)([^0.,'#%].*)?"
DATEFORMAT_RE = "|".join([r"{}+".format(switch) for switch in "yYmMdDhHsS"] + [r"am/pm", r"AM/PM"])
DATEFORMAT_MAP = {
    "M": "{d.month}",
    "MM": "%m",
    "MMM": "%b",
    "MMMM": "%B",
    "d": "{d.day}",
    "dd": "%d",
    "ddd": "%a",
    "dddd": "%A",
    "D": "{d.day}",
    "DD": "%d",
    "DDD": "%a",
    "DDDD": "%A",
    "yy": "%y",
    "yyyy": "%Y",
    "YY": "%y",
    "YYYY": "%Y",
    "h": "{hour12}",
    "hh": "%I",
    "H": "{d.hour}",
    "HH": "%H",
    "m": "{d.minute}",
    "mm": "%M",
    "s": "{d.second}",
    "ss": "%S",
    "am/pm": "%p",
    "AM/PM": "%p",
}

TAGS_WITH_ID = {
    'wp:docPr': {'name': 'Picture {id}'}
}

MAKE_TESTS_HAPPY = True

class NextRecord(Exception):
    pass

class SkipRecord(Exception):
    pass

class MergeField(object):
    """
    Base MergeField class

    it contains the field name, and a method to return a list of elements (runs) given the data dictionary 
    """

    def __init__(self, parent, name="", key=None, instr=None, instr_tokens=None, nested=False, elements=None, ignore_elements=None):
        """ Inits the MergeField class
        
        Args:
            parent: The parent element of the MergeField in the tree.
            idx: The idx of the MergeField in the parent.
            name: The name of the field, if applicable"""
        self.parent = parent
        self.nested = nested
        self.key = key # the key of this MergeField to be able to identify it. It is used as the name in the replaced MergeField element
        # the list of elements to add when merging
        self._elements_to_add = [] if elements is None else elements
        # the list of elements to ignore when merging (after and including "separate" fldCharType)
        self._elements_to_ignore = [] if ignore_elements is None else ignore_elements
        self.instr = instr
        self.instr_tokens = instr_tokens
        self.filled_elements = []
        self.filled_value = None
        self.name = name
        if not name and instr_tokens[1:]:
            self.name = instr_tokens[1]

        self._parse_instruction()

    def _parse_instruction(self):
        # # we keep only the first node, and will add the text to this node
        self._elements_to_ignore.extend(self._elements_to_add[1:])
        del self._elements_to_add[1:]
        first_elem = self._elements_to_add[0]
        if first_elem.tag == '{%(w)s}fldSimple' % NAMESPACES:
            # workaround for simple fields
            elem_to_add = first_elem.find('{%(w)s}r' % NAMESPACES)
            if MAKE_TESTS_HAPPY:
                elem_to_add.clear()
            self.parent.replace(first_elem, elem_to_add)
            self._elements_to_add[0] = elem_to_add
    
    def _format(self, value):
        options = self.instr_tokens[2:]
        while options:
            flag = options[0][0:2]
            if not flag:
                options = options[1:]
                continue
            if options[0][2:]: # no space after the flag
                option = options[0][2:]
                options = options[1:]
            else:
                option = options[1]
                options = options[2:]
            if flag in ('\\b', '\\f'):
                value = self._format_bf(value, flag, option)
            if flag in ('\\#'):
                value = self._format_number(value, flag, option)
            if flag in ('\\@'):
                value = self._format_date(value, flag, option)
            if flag in ('\\*'):
                value = self._format_text(value, flag, option)

        if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
            # TODO format the date according to the locale -- set the locale
            date_formats = []
            if hasattr(value, 'month'):
                date_formats.append('%x')
            if hasattr(value, 'hour'):
                date_formats.append('%X')
            value = value.strftime(" ".join(date_formats))

        return value

    def _format_bf(self, value, flag, option):
        # print("<{}><{}>".format(value, type(value)))
        if value:
            if flag == '\\b':
                return option + str(value)
            return str(value) + option
        # no value no text
        return None

    def _format_text(self, value, flag, option):
        option = option.lower()
        if option == 'caps':
            return str(value).title()
        if option == 'firstcap':
            return str(value).capitalize()
        if option == 'upper':
            return str(value).upper()
        if option == 'lower':
            return str(value).lower()

        return value

    def _format_number(self, value, flag, option):
        format_match = re.match(NUMBERFORMAT_RE, option)
        if value is None:
            value = 0
        if format_match is None:
            warnings.warn("Non conforming number format <{}>".format(option))
            return value
        format_prefix = format_match.group(1) or ""
        format_number = format_match.group(2)
        format_suffix = format_match.group(3) or ""
        if format_number[0] == 'P':
            return "{{}}{{:.{}%}}{{}}".format(
                int(format_number[1:])).format(
                    format_prefix,
                    value,
                    format_suffix)
        if format_number[0] == 'N':
            return "{{}}{{:.{}f}}{{}}".format(
                int(format_number[1:])).format(
                    format_prefix,
                    value,
                    format_suffix)
        if format_number[-1] == '%':
            return "{}{:.0%}{}".format(
                    format_prefix,
                    value,
                    format_suffix)
        thousand_info = [
            ('_', thousand_char)
            for thousand_char in "',"
            if thousand_char in format_number] + [('', '')]
        thousand_flag, thousand_char = thousand_info[0]
        format_number = format_number.replace(',', '')
        digits, decimals = (format_number.split('.') + [''])[0:2]
        zero_digits = len(digits.replace('#', ''))
        zero_decimals = len(decimals.replace('#', ''))
        len_decimals_plus_dot = 0 if not decimals else 1 + len(decimals)
        number_format_text = "{{}}{{:{zero_digits}{thousand_flag}{decimals}f}}{{}}".format(
            thousand_flag=thousand_flag,
            zero_digits="0>{}".format(zero_digits + len_decimals_plus_dot) if zero_digits > 1 else "",
            decimals=".{}".format(len(decimals)))
        # print(self.name, "<", option, ">", number_format_text)
        try:
            result = number_format_text.format(
                        format_prefix,
                        value,
                        format_suffix)
            if thousand_flag:
                result = result.replace(thousand_flag, thousand_char)
            return result
        except Exception as e:
            raise ValueError("Invalid number format <{}> with error <{}>".format(number_format_text, e))

    def _format_date(self, value, flag, option):

        if value is None:
            return ''

        if not isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
            return str(value)

        # set the locale to be used for time
        # more checking needed before activating
        # locale.setlocale(locale.LC_TIME, "")
        fmt = re.sub(DATEFORMAT_RE, lambda x:DATEFORMAT_MAP[x[0]], option)
        fmt_args = {'d': value}
        if hasattr(value, 'hour'):
            fmt_args['hour12'] = value.hour % 12
        fmt = fmt.format(**fmt_args)
        value = value.strftime(fmt)
        # warnings.warn("Date formatting not yet implemented <{}>".format(option))
        return value

    def fill_data(self, merge_data, row):
        self.filled_elements = []
        value = row.get(self.name, "UNKNOWN({})".format(self.instr))
        try:
            value = self._format(value)
        except Exception as e:
            warnings.warn("Invalid formatting for field <{}> with error <{}>".format(self.instr, e))
            # raise

        self.filled_value = value

        if value is None:
            # no elements should be filled ?
            # or empty text ?
            if MAKE_TESTS_HAPPY:
                value = ""
            else:
                return
        
        elem = deepcopy(self._elements_to_add[0])
        for child in elem.xpath('w:instrText', namespaces=NAMESPACES):
            elem.remove(child)
        for child in elem.xpath('w:t', namespaces=NAMESPACES):
            elem.remove(child)

        text_parts = str(value).replace('\r', '').split('\n')
        elem.append(self._make_text(text_parts[0]))
        for text_part in text_parts[1:]:
            elem.append(self._make_br())
            elem.append(self._make_text(text_part))

        self.filled_elements.append(elem)

    def _make_br(self):
        return Element('{%(w)s}br' % NAMESPACES)
    
    def _make_text(self, text):
        if self.nested:
            text_node = Element('{%(w)s}instrText' % NAMESPACES)
            text_node.set("{%(xml)s}space" % NAMESPACES, "preserve")
        else:
            text_node = Element('{%(w)s}t' % NAMESPACES)

        text_node.text = text
        return text_node

    def insert_into_tree(self):
        """ inserts a MergeField element in the original tree at the right position
        
        """
        # Make sure ALL elements from the original tree are removed except for the first useful, that we will replace
        for subelem in self._elements_to_ignore:
            self.parent.remove(subelem)
        for subelem in self._elements_to_add[1:]:
            self.parent.remove(subelem)

        replacement_element = Element("MergeField", merge_key=self.key, name=self.name)
        self.parent.replace(self._elements_to_add[0], replacement_element)
        return replacement_element

class NextField(MergeField):

    def fill_data(self, merge_data, row):
        raise NextRecord()

class MergeData(object):

    """ prepare the MergeField objects and the data """

    FIELD_CLASSES = {
        "MERGEFIELD": MergeField,
        "NEXT": NextField,
    }

    def __init__(self, remove_empty_tables=False):
        self._merge_field_map = {} # merge_field.key: MergeField()
        self._merge_field_next_id = 0
        self.duplicate_id_map = {} # tag: {'max': max_id, 'ids': set(existing_ids)}
        self.has_nested_fields = False
        self.remove_empty_tables = remove_empty_tables
        self._rows = None
        self._current_index = None

    def start_merge(self, replacements):
        assert self._rows is None, "merge already started"
        self._rows = replacements
        return self.next_row()
    
    def next_row(self):
        assert self._rows is not None, "merge not yet started"

        if self._current_index is None:
            self._current_index = 0
        else:
            self._current_index += 1
        
        if self._current_index < len(self._rows):
            return self._rows[self._current_index]
    
    def is_first(self):
        return self._current_index == 0
    
    def get_new_element_id(self, element):
        """ Returns None if the existing id is new otherwise a new id """
        tag = element.tag
        id = element.get('id')
        if id is None:
            return None
        id = int(id)
        id_data = self.duplicate_id_map.setdefault(tag, {'max': id, 'values': set()})
        if id in id_data['values']:
            # it already exists
            id = id_data['max'] + 1
            id_data['values'].add(id)
            id_data['max'] = id
            return str(id)

        id_data['values'].add(id)
        id_data['max'] = max(id, id_data['max'])
        return None

    def get_merge_fields(self, key):
        merge_obj = self.get_field_obj(key)
        if merge_obj.name:
            yield merge_obj.name

    def get_instr_text(self, elements, recursive=False):
        return "".join([
            text
            for elem in elements
            for text in elem.xpath('w:instrText/text()', namespaces=NAMESPACES) + [
                "{{{}}}".format(obj_name) if not recursive else self.get_field_obj(obj_name).instr for obj_name in elem.xpath('@merge_key')]
        ])

    @classmethod
    def _get_instr_tokens(cls, instr):
        s = shlex.shlex(instr, posix=True)
        s.whitespace_split = True
        s.commenters = ''
        s.escape = ''
        return s

    @classmethod
    def _get_field_type(cls, instr):
        s = shlex.split(instr, posix=False)
        return s[0], s[1:]

    def make_data_field(self, parent, key=None, nested=False, instr=None, elements=None, **kwargs):
        """ MergeField factory method """
        if key is None:
            key = self._get_next_key()

        instr = instr or self.get_instr_text(elements)
        field_type, rest = self._get_field_type(instr)
        field_class = self.FIELD_CLASSES.get(field_type)
        if field_class is None:
            # ignore the field
            # print("ignore field", instr)
            return None

        try:
            tokens = list(self._get_instr_tokens(instr))
        except ValueError as e:
            tokens = [field_type] + list(map(lambda part:part.replace('"', ''), rest))
            warnings.warn("Invalid field description <{}> near: <{}>".format(str(e), instr))

        # print("make data object", field_class, instr, len(elements), len(kwargs.get('ignore_elements', [])))
        field_obj = field_class(
            parent,
            key=key,
            instr=instr,
            nested=nested,
            instr_tokens=tokens,
            elements=elements,
            **kwargs
        )

        assert key not in self._merge_field_map
        if nested:
            self.has_nested_fields = True
        self._merge_field_map[key] = field_obj
        return field_obj

    def get_field_obj(self, key):
        return self._merge_field_map[key]

    def mark_field_as_nested(self, key, nested=True):
        if nested:
            self.has_nested_fields = True
        self.get_field_obj(key).nested = nested

    def _get_next_key(self):
        key = "field_{}".format(self._merge_field_next_id)
        self._merge_field_next_id += 1
        return key

    def replace(self, body, row):
        """ replaces in the body xml tree the MergeField elements with values from the row """
        all_tables = {
            key: value
            for key, value in row.items()
            if isinstance(value, list)
            }

        for anchor, table_rows in all_tables.items():
            self.replace_table_rows(body, anchor, table_rows)

        merge_fields = body.findall('.//MergeField')
        for field_element in merge_fields:
            try:
                filled_field = self.fill_field(field_element, row)
                if filled_field:
                    for text_element in reversed(filled_field.filled_elements):
                        field_element.addnext(text_element)
                    field_element.getparent().remove(field_element)
            except NextRecord:
                field_element.getparent().remove(field_element)
                row = self.next_row()
    
    def replace_table_rows(self, body, anchor, rows):
        """ replace the rows of a table with the values from the rows list """
        table, idx, template = self.__find_row_anchor(body, anchor)
        if table is not None:
            if len(rows) > 0:
                del table[idx]
                for i, row_data in enumerate(rows):
                    row = deepcopy(template)
                    self.replace(row, row_data)
                    table.insert(idx + i, row)
            else:
                # if there is no data for a given table
                # we check whether table needs to be removed
                if self.remove_empty_tables:
                    parent = table.getparent()
                    parent.remove(table)

    def __find_row_anchor(self, body, field):
        for table in body.findall('.//{%(w)s}tbl' % NAMESPACES):
            for idx, row in enumerate(table):
                if row.find('.//MergeField[@name="%s"]' % field) is not None:
                    return table, idx, row
        return None, None, None
    
    def fill_field(self, field_element, row):
        """" fills the corresponding MergeField python object with data from row """
        if field_element.get('name') and (row is None or field_element.get('name') not in row):
            return None
        field_key = field_element.get('merge_key')
        field_obj = self._merge_field_map[field_key]
        field_obj.fill_data(self, row)
        return field_obj

class MergeDocument(object):

    """ prepare and merge one document 
    
        helper class to handle the actual merging of one document
        It prepares the body, sections, separators
    """

    def __init__(self, root, separator):
        self.root = root
        # self.sep_type = sep_type
        # self.sep_class = sep_class
        # if sep_class == 'section':
        #     self._set_section_type()
        
        self._last_section = None # saving the last section to add it at the end
        self._body = None # the document body, where all the documents are appended
        self._body_copy = None # a deep copy of the original body without ending section
        self._current_body = None # the current document body where all the changes are merged
        self._prepare_data(separator)

    def _prepare_data(self, separator):

        if separator not in VALID_SEPARATORS:
            raise ValueError("Invalid separator argument")
        sep_type, sep_class = separator.split("_")

        # TODO why setting the type only to the first section and not to the last section ?

        if sep_class == 'section':
            #FINDING FIRST SECTION OF THE DOCUMENT
            first_section = self.root.find("w:body/w:p/w:pPr/w:sectPr", namespaces=NAMESPACES)
            if first_section is None:
                first_section = self.root.find("w:body/w:sectPr", namespaces=NAMESPACES)
        
            type_element = first_section.find("w:type", namespaces=NAMESPACES)
            
            if MAKE_TESTS_HAPPY:
                if type_element is not None:
                    first_section.remove(type_element)
                    type_element = None

            if type_element is None:
                type_element = etree.SubElement(first_section, '{%(w)s}type' % NAMESPACES)

            type_element.set('{%(w)s}val' % NAMESPACES, sep_type)

        #FINDING LAST SECTION OF THE DOCUMENT
        self._last_section = self.root.find("w:body/w:sectPr", namespaces=NAMESPACES)
    
        self._body = self._last_section.getparent()
        self._body.remove(self._last_section)

        self._body_copy = deepcopy(self._body)

        #EMPTY THE BODY - PREPARE TO FILL IT WITH DATA
        self._body.clear()

        self._separator = etree.Element('{%(w)s}p'  % NAMESPACES)

        if sep_class == 'section':
            pPr = etree.SubElement(self._separator, '{%(w)s}pPr'  % NAMESPACES)
            pPr.append(deepcopy(self._last_section))
        elif sep_class == 'break':
            r = etree.SubElement(self._separator, '{%(w)s}r'  % NAMESPACES)
            nbreak = etree.SubElement(r, '{%(w)s}br' % NAMESPACES)
            nbreak.set('{%(w)s}type' % NAMESPACES, sep_type)
    
    def prepare(self, merge_data, first=False):
        """ prepares the current body for the merge """
        assert self._current_body is None
        # add separator if not the first document
        if not first:
            self._body.append(deepcopy(self._separator))
        self._current_body = deepcopy(self._body_copy)
        for tag, attr_gen in TAGS_WITH_ID.items():
            for elem in self._current_body.xpath('//{}'.format(tag), namespaces=NAMESPACES):
                self._fix_id(merge_data, elem, attr_gen)

    def merge(self, merge_data, row, first=False):
        """ Merges one row into the current prepared body """
        
        merge_data.replace(self._current_body, row)

    def _fix_id(self, merge_data, element, attr_gen):
        new_id = merge_data.get_new_element_id(element)
        if new_id is not None:
            element.attrib['id'] = new_id
            for attr_name, attr_value in attr_gen.items():
                element.attrib[attr_name] = attr_value.format(id=new_id)

    def finish(self, abort=False):
        """ finishes the current body by saving it into the main body or into a file (future feature) """

        if abort: # for skipping the record
            self._current_body = None

        if self._current_body is not None:
            for child in self._current_body:
                self._body.append(child)
            self._current_body = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            self.finish(True)
            self._body.append(self._last_section)

class MailMerge(object):

    """
    MailMerge class to write an output docx document by merging data rows to a template

    The class uses the builtin MergeFields in Word. There are two kind of data fields, simple and complex.
    http://officeopenxml.com/WPfields.php
    The MERGEFIELD can have MERGEFORMAT
    MERGEFIELD can be nested inside other "complex" fields, in which case those fields should be updated in the saved docx

    MailMerge implements this by finding all Fields and replacing them with placeholder Elements of type
    MergeElement

    Those MergeElement elements will then be replaced for each run with a list of elements containing run elements with texts.
    The MergeElement value (list of run Elements) should be computed recursively for the inner MergeElements

    The IF MergeElement has the condition MergeCondition, and two lists of elements, one for true and one for false

    """

    def __init__(self, file, remove_empty_tables=False, auto_update_fields_on_open="no"):
        """ 
        auto_update_fields_on_open : no, auto, always - auto = only when needed
        """
        self.zip = ZipFile(file)
        self.parts = {} # part: ElementTree
        self.settings = None
        self._settings_info = None
        self.merge_data = MergeData(remove_empty_tables=remove_empty_tables)
        self.remove_empty_tables = remove_empty_tables
        self.auto_update_fields_on_open = auto_update_fields_on_open

        try:
            self.__fill_parts()

            for part in self.parts.values():
                self.__fill_simple_fields(part)
                self.__fill_complex_fields(part)

            # Remove mail merge settings to avoid error messages when opening document in Winword
            self.__fix_settings()
        except:
            self.zip.close()
            raise


    def __setattr__(self, __name, __value):
        super(MailMerge, self).__setattr__(__name, __value)
        if __name == 'remove_empty_tables':
            self.merge_data.remove_empty_tables = __value

    def __fill_parts(self):
        content_types = etree.parse(self.zip.open('[Content_Types].xml'))
        for file in content_types.findall('{%(ct)s}Override' % NAMESPACES):
            type = file.attrib['ContentType' % NAMESPACES]
            if type in CONTENT_TYPES_PARTS:
                zi, self.parts[zi] = self.__get_tree_of_file(file)
            elif type == CONTENT_TYPE_SETTINGS:
                self._settings_info, self.settings = self.__get_tree_of_file(file)
    
    def __fill_simple_fields(self, part):
        for fld_simple_elem in part.findall('.//{%(w)s}fldSimple' % NAMESPACES):
            merge_field_obj = self.merge_data.make_data_field(
                fld_simple_elem.getparent(), instr=fld_simple_elem.get('{%(w)s}instr' % NAMESPACES), elements=[fld_simple_elem])
            if merge_field_obj:
                merge_field_obj.insert_into_tree()

    def __get_next_element(self, current_element):
        """ returns the next element of a complex field """
        next_element = current_element.getnext()
        current_paragraph = current_element.getparent()
        # we search through paragraphs for the next <w:r> element
        while next_element is None:
            current_paragraph = current_paragraph.getnext()
            if current_paragraph is None:
                return None, None, None
            next_element = current_paragraph.find('w:r', namespaces=NAMESPACES)
        
        # print(''.join(next_element.xpath('w:instrText/text()', namespaces=NAMESPACES)))
        field_char_subelem = next_element.find('w:fldChar', namespaces=NAMESPACES)
        if field_char_subelem is None:
            return next_element, None, None

        return next_element, field_char_subelem, field_char_subelem.xpath('@w:fldCharType', namespaces=NAMESPACES)[0]

    def _pull_next_merge_field(self, elements_of_type_begin, nested=False):

        assert (elements_of_type_begin)
        current_element = elements_of_type_begin.pop(0)
        parent_element = current_element.getparent()
        good_elements = []
        ignore_elements = [current_element]
        current_element_list = good_elements
        field_char_type = None

        # print('>>>>>>>')
        while field_char_type != 'end':
            # find next sibling
            next_element, field_char_subelem, field_char_type = \
                self.__get_next_element(current_element)
            if next_element is None:
                
                instr_text = self.merge_data.get_instr_text(good_elements, recursive=True)
                raise ValueError("begin without end near:" + instr_text)

            if field_char_type == 'begin':
                # nested elements
                assert(elements_of_type_begin[0] is next_element)
                merge_field_sub_obj, next_element = self._pull_next_merge_field(elements_of_type_begin, nested=True)
                if merge_field_sub_obj:
                    next_element = merge_field_sub_obj.insert_into_tree()
                # print("current list is ignore", current_element_list is ignore_elements)
                # print("<<<<< #####", etree.tostring(next_element))
            elif field_char_type == 'separate':
                current_element_list = ignore_elements
            elif next_element.tag == 'MergeField':
                # we have a nested simple Field - mark it as nested
                self.merge_data.mark_field_as_nested(next_element.get('merge_key'))

            current_element_list.append(next_element)
            current_element = next_element
        
        # print('<<<<<<<', len(good_elements), len(ignore_elements))
        merge_obj = self.merge_data.make_data_field(
            parent_element,
            nested=nested,
            elements=good_elements,
            ignore_elements=ignore_elements)
        return merge_obj, current_element

    def __fill_complex_fields(self, part):
        """ finds all begin fields and then builds the MergeField objects and inserts the replacement Elements in the tree """
        elements_of_type_begin = list(part.findall('.//{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="begin"]/..' % NAMESPACES))
        while elements_of_type_begin:
            merge_field_obj, _ = self._pull_next_merge_field(elements_of_type_begin)
            if merge_field_obj:
                # print(merge_field_obj.instr)
                merge_field_obj.insert_into_tree()

    def __fix_settings(self):

        if self.settings:
            settings_root = self.settings.getroot()
            mail_merge = settings_root.find('{%(w)s}mailMerge' % NAMESPACES)
            if mail_merge is not None:
                settings_root.remove(mail_merge)
            
            add_update_fields_setting = (
                self.auto_update_fields_on_open == "auto" and self.merge_data.has_nested_fields
                or self.auto_update_fields_on_open == "always"
            )
            if add_update_fields_setting:
                update_fields_elem = settings_root.find('{%(w)s}updateFields' % NAMESPACES)
                if not update_fields_elem:
                    update_fields_elem = etree.SubElement(settings_root, '{%(w)s}updateFields' % NAMESPACES)
                update_fields_elem.set('{%(w)s}val' % NAMESPACES, "true")

    def __get_tree_of_file(self, file):
        fn = file.attrib['PartName' % NAMESPACES].split('/', 1)[1]
        zi = self.zip.getinfo(fn)
        return zi, etree.parse(self.zip.open(zi))

    def write(self, file, empty_value=''):
        if empty_value is not None:
            self.merge(**{
                field: empty_value
                for field in self.get_merge_fields()
            })

        with ZipFile(file, 'w', ZIP_DEFLATED) as output:
            for zi in self.zip.filelist:
                if zi in self.parts:
                    xml = etree.tostring(self.parts[zi].getroot(), encoding='UTF-8', xml_declaration=True)
                    output.writestr(zi.filename, xml)
                elif zi == self._settings_info:
                    xml = etree.tostring(self.settings.getroot(), encoding='UTF-8', xml_declaration=True)
                    output.writestr(zi.filename, xml)
                else:
                    output.writestr(zi.filename, self.zip.read(zi))

    def get_merge_fields(self, parts=None):
        if not parts:
            parts = self.parts.values()

        fields = set()
        for part in parts:
            for mf in part.findall('.//MergeField'):
                for name in self.merge_data.get_merge_fields(mf.attrib['merge_key']):
                    fields.add(name)
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
        assert replacements, "empty data"
        #TYPE PARAM CONTROL AND SPLIT
        
        #GET ROOT - WORK WITH DOCUMENT
        for part in self.parts.values():
            root = part.getroot()
            tag = root.tag

            # ignore header and footer ?
            if tag == '{%(w)s}ftr' % NAMESPACES or tag == '{%(w)s}hdr' % NAMESPACES:
                continue

            with MergeDocument(root, separator) as merge_doc:
                row = self.merge_data.start_merge(replacements)
                while row is not None:
                    merge_doc.prepare(self.merge_data, first=self.merge_data.is_first())
                    try:
                        merge_doc.merge(self.merge_data, row)
                        merge_doc.finish()
                    except SkipRecord:
                        merge_doc.finish(abort=True)
                    row = self.merge_data.next_row()

    def merge_pages(self, replacements):
         """
         Deprecated method.
         """
         warnings.warn("merge_pages has been deprecated in favour of merge_templates",
                      category=DeprecationWarning,
                      stacklevel=2)         
         self.merge_templates(replacements, "page_break")

    def merge(self, parts=None, **replacements):
        if not parts:
            parts = self.parts.values()

        for part in parts:
            self.merge_data.replace(part, replacements)

    def merge_rows(self, anchor, rows):
        """ anchor is one of the fields in the table """

        for part in self.parts.values():
            self.merge_data.replace_table_rows(part, anchor, rows)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        if self.zip is not None:
            try:
                self.zip.close()
            finally:
                self.zip = None
