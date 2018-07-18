from unittest import TestCase
from datetime import datetime
from mailmerge import MailMerge


class TestEval(TestCase):

    def setUp(self):
        self.today = datetime(2018, 7, 1, 10, 12, 54, 0)

    def test_year(self):
        self.assertEqual(MailMerge.eval_strftime(self.today, 'y'), '2018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'Y'), '2018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'yy'), '18')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'YY'), '18')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'yyyy'), '2018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'YYYY'), '2018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'yyyyy'), '20182018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'YYYYY'), '20182018')

    def test_month(self):
        self.assertEqual(MailMerge.eval_strftime(self.today, 'M'), '7')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'MM'), '07')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'MMMM'), self.today.strftime('%B'))

    def test_days(self):
        self.assertEqual(MailMerge.eval_strftime(self.today, 'd'), '1')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'dd'), '01')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'dddd'), self.today.strftime('%A'))

    def test_strftime_none(self):
        self.assertEqual(MailMerge.eval_strftime(None, 'd'), '')

    def test_strftime_str(self):
        self.assertEqual(MailMerge.eval_strftime('foo', 'd'), 'foo')

    def test_eval(self):
        self.assertEqual(MailMerge.eval_star(self.today), self.today.strftime('%x %X'))
        self.assertEqual(MailMerge.eval_star(self.today.date()), self.today.strftime('%x'))
        self.assertEqual(MailMerge.eval_star(self.today.time()), self.today.strftime('%X'))

    def test_eval_none(self):
        self.assertEqual(MailMerge.eval(None, 'MAILMERGE Foo \\@ "y-MM-dd" MERGEFORMAT'), '')

    def test_parse_code(self):
        self.assertEqual(MailMerge.eval(self.today, 'MAILMERGE Foo \\* MERGEFORMAT'), self.today.strftime('%x %X'))
        self.assertEqual(MailMerge.eval(self.today, 'MAILMERGE Foo'), self.today.strftime('%x %X'))
        self.assertEqual(MailMerge.eval(self.today, 'MAILMERGE Foo \\@ "y-MM-dd" MERGEFORMAT'), '2018-07-01')
