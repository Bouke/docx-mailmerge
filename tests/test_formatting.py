import unittest
import tempfile
import warnings
import datetime
from os import path
from lxml import etree

from mailmerge import MailMerge, NAMESPACES
from tests.utils import EtreeMixin, get_document_body_part

class FormattingTest(EtreeMixin, unittest.TestCase):

    def setUp(self):
        self.open_docx(path.join(path.dirname(__file__), 'test_one_simple_field.docx'))
        zi, part = self.get_part()
        self.replacement_parts = {zi: part}
        self.simple_merge_field = part.getroot().find('.//{%(w)s}fldSimple' % NAMESPACES)
    
    def _test_formats(self, flag, format_tests):
        for formatting, value_list in format_tests.items():
            rows = [{'fieldname': value} for value, _ in value_list]
            # print(formatting)

            instr = 'MERGEFIELD fieldname {} "{}"'.format(flag, formatting)
            self.simple_merge_field.set('{%(w)s}instr' % NAMESPACES, instr)
            with MailMerge(self.get_new_docx(self.replacement_parts)) as document:
                self.assertEqual(document.get_merge_fields(), {'fieldname'})
                document.merge_templates(rows, 'page_break')

                with tempfile.TemporaryFile() as outfile:
                    document.write(outfile)

            output_fields = get_document_body_part(document).getroot().xpath('.//w:r/w:t/text()', namespaces=NAMESPACES)
            self.assertEqual(output_fields, [output_value for _, output_value in value_list])

    
    def test_number(self):
        self._test_formats(
            '\\#',
            {
                "0.00": [
                    (25, "25.00"),
                    (0, "0.00"),
                    (345.12314, "345.12"),
                    (None, "0.00")
                    ],
                "#,#00": [
                    (25, "25"),
                    (0, "00"),
                    (32345.12314, "32,345")
                    ],
                "#,###.##": [
                    (0, "0.00"),
                    (23423, "23,423.00")
                ],
                "#'###.##": [
                    (0, "0.00"),
                    (23423, "23'423.00")
                ],
                "#,##0.00": [
                    (0, "0.00"),
                    (23423, "23,423.00")
                ],
                "N3": [
                    (25, "25.000"),
                    (0, "0.000"),
                    (345.12314, "345.123")
                    ],
                "P3": [
                    (25, "2500.000%"),
                    (0, "0.000%"),
                    (345.12314, "34512.314%")
                    ],
                "##%": [
                    (0.25, "25%"),
                    (0, "0%"),
                    (0.034512314, "3%")
                    ],
                # "#.##": [
                #     (0.25, ".25"),
                #     (5, "5."),
                #     ],
            })

    def test_date(self):
        datetime_value = datetime.datetime(2022, 4, 12, 15, 10, 59)
        date_value = datetime_value.date()
        time_value = datetime_value.time()
        return
        # test date values
        for value in [datetime_value, date_value]:
            self._test_formats(
                '\\@', 
                {  
                    'M': [(value, '')],
                    'MM': [(value, '')],
                    'MMM': [(value, '')],
                    'MMMM': [(value, '')],
                    'd': [(value, '')],
                    'dd': [(value, '')],
                    'ddd': [(value, '')],
                    'dddd': [(value, '')],
                    'D': [(value, '')],
                    'DD': [(value, '')],
                    'DDD': [(value, '')],
                    'DDDD': [(value, '')],
                    'yy': [(value, '')],
                    'yyyy': [(value, '')],
                    'YY': [(value, '')],
                    'YYYY': [(value, '')],
                }
            )

        # test time values
        for value in [datetime_value, time_value]:
            self._test_formats(
                '\\@', 
                {  
                    'h': [(value, '')],
                    'hh': [(value, '')],
                    'H': [(value, '')],
                    'HH': [(value, '')],
                    'm': [(value, '')],
                    'mm': [(value, '')],
                    's': [(value, '')],
                    'ss': [(value, '')],
                    'am/pm': [(value, '')],
                    'AM/PM': [(value, '')],
                }
            )
        
        # special cases
        # TODO add empty value, null value, string value
        # TODO also add datetime/date/time values with no format with different locales

    def test_text(self):
        self._test_formats(
            '\\*',
            {
                "Caps": [
                    ("handsome law group", "Handsome Law Group"),
                    ("HANDSOME  LAW GROUP", "Handsome  Law Group")
                ],
                "FirstCap": [
                    ("handsome law group", "Handsome law group"),
                ],
                "Upper": [
                    ("Handsome Law Group", "HANDSOME LAW GROUP"),
                ],
                "Lower": [
                    ("Handsome Law Group", "handsome law group"),
                ],
                "Invalid": [
                    ("handsome Law Group", "handsome Law Group"),
                ]
            })

    def test_text_before(self):
        self._test_formats(
            '\\b',
            {
                " - ": [
                    ("handsome law group", " - handsome law group"),
                    ("", "")
                ]
            }
        )

    def test_text_afterward(self):
        self._test_formats(
            '\\f',
            {
                " - ": [
                    ("handsome law group", "handsome law group - "),
                    ("", "")
                ]
            }
        )

    def test_text_b_and_f(self):
        self._test_formats(
            '\\b',
            {
                "(\" \\f \")": [
                    (2, "(2)"),
                    ("", "")
                ]
            }
        )

    def test_invalid_formatting_syntax(self):
        with warnings.catch_warnings(record=True) as warning_list:
            self._test_formats(
                '\\b',
                {
                    "(\" \\f": [
                        (2, "2"),
                    ]
                }
            )
            self.assertTrue(warning_list)

    def test_invalid_formatting(self):
        with warnings.catch_warnings(record=True) as warning_list:
            self._test_formats(
                '\\b',
                {
                    "(\" \\f": [
                        (2, "2"),
                    ]
                }
            )
            self.assertTrue(warning_list)

    def tearDown(self):
        self.docx_zipfile.close()
