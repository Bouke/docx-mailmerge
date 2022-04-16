import unittest
import tempfile
import warnings
import datetime
import locale
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
            self.assertEqual(output_fields, [output_value for _, output_value in value_list], "Format <{} {}>".format(flag, formatting))

    
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
        datetime_value = datetime.datetime(2022, 3, 9, 17, 7, 8)
        date_value = datetime_value.date()
        time_value = datetime_value.time()
        # time12_value = datetime_value - datetime.timedelta(hours=12)
        locale.setlocale(locale.LC_TIME, 'en_US')

        # https://support.microsoft.com/en-us/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c#ID0EBBD=Date-Time_format_switch_(\@)

        # test date values
        for value in [datetime_value, date_value]:
            self._test_formats(
                '\\@', 
                {  
                    # Month (M)
                    # The letter M must be uppercase to distinguish months from minutes.
                    # M    This format item displays the month as a number without a leading 0 (zero) for single-digit months. For example, July is 7.
                    # MM    This format item displays the month as a number with a leading 0 (zero) for single-digit months. For example, July is 07.
                    # MMM    This format item displays the month as a three-letter abbreviation. For example, July is Jul.
                    # MMMM    This format item displays the month as its full name.                    
                    'M': [(value, '3')],
                    'MM': [(value, '03')],
                    'MMM': [(value, 'Mar')],
                    'MMMM': [(value, 'March')],
                    # Day (d)
                    # The letter d displays the day of the month or the day of the week. The letter d can be either uppercase or lowercase.
                    # d    This format item displays the day of the week or month as a number without a leading 0 (zero) for single-digit days. For example, the sixth day of the month is displayed as 6.
                    # dd    This format item displays the day of the week or month as a number with a leading 0 (zero) for single-digit days. For example, the sixth day of the month is displayed as 06.
                    # ddd    This format item displays the day of the week or month as a three-letter abbreviation. For example, Tuesday is displayed as Tue.
                    # dddd    This format item displays the day of the week as its full name.
                    'd': [(value, '9')],
                    'dd': [(value, '09')],
                    'ddd': [(value, 'Wed')],
                    'dddd': [(value, 'Wednesday')],
                    'D': [(value, '9')],
                    'DD': [(value, '09')],
                    'DDD': [(value, 'Wed')],
                    'DDDD': [(value, 'Wednesday')],
                    # Year (y)
                    # The letter y displays the year as two or four digits. The letter y can be either uppercase or lowercase.
                    # yy    This format item displays the year as two digits with a leading 0 (zero) for years 01 through 09. For example, 1999 is displayed as 99, and 2006 is displayed as 06.
                    # yyyy    This format item displays the year as four digits.
                    'yy': [(value, '22')],
                    'yyyy': [(value, '2022')],
                    'YY': [(value, '22')],
                    'YYYY': [(value, '2022')],
                    # combined
                    'YYYYMMDD': [(value, '20220309')]
                }
            )

        # test time values 17:07:08
        for value in [datetime_value, time_value]:
            self._test_formats(
                '\\@', 
                {
                    'h': [(value, '5')],
                    'hh': [(value, '05')],
                    'H': [(value, '17')],
                    'HH': [(value, '17')],
                    'm': [(value, '7')],
                    'mm': [(value, '07')],
                    's': [(value, '8')],
                    'ss': [(value, '08')],
                    'am/pm': [(value, 'PM')],
                    'AM/PM': [(value, 'PM')],
                }
            )

        self._test_formats('', {'': [(datetime_value, '03/09/2022 17:07:08')]})
        self._test_formats('', {'': [(date_value, '03/09/2022')]})
        self._test_formats('', {'': [(time_value, '17:07:08')]})

        locale.setlocale(locale.LC_TIME, 'de_DE')
        self._test_formats(
            '\\@',
            {
                'MMM': [(datetime_value, 'Mär')],
                'am/pm': [(datetime_value, 'pm')],
            }
        )

        self._test_formats('', {'': [(datetime_value, '09.03.2022 17:07:08')]})
        self._test_formats('', {'': [(date_value, '09.03.2022')]})
        self._test_formats('', {'': [(time_value, '17:07:08')]})

        # special cases
        # TODO add empty value, null value, string value
        self._test_formats(
            '\\@',
            {
                'MMM': [(None, '')],
                'am/pm': [('2022-03-09', '2022-03-09')],
            }
        )
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
                    '(" \\f': [
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
