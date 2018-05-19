# encoding: utf-8

import unittest
import sys
sys.path.append('..')

from workbookwithheader import *


class WorkbookWithHeaderTest(unittest.TestCase):

    workbook_path = os.path.abspath(os.path.join(os.path.dirname(os.getcwd()), 'samp', 'Sample_Spread_with_Header.xlsx'))

    def setUp(self):
        print('Setup for WorkbookWithHeaderTest')

    def tearDown(self):
        print('Teardown for WorkbookWithHeaderTest')

    def workbook_load_tests(self):

        workbook_obj = WorkbookWithHeader()
        workbook_obj.load_workbook(self.workbook_path)

        self.assertIn('KeysInHeader', workbook_obj.get_tab_list())
        self.assertIn('KeysInRows', workbook_obj.get_tab_list())
        self.assertIn('Configs', workbook_obj.get_tab_list())

        for worksheet_name in workbook_obj.get_tab_list():
            self.assertIsInstance(workbook_obj.get_worksheet_by_name(worksheet_name), WorksheetWithHeader)
            self.assertEqual(worksheet_name, workbook_obj.get_worksheet_by_name(worksheet_name).get_worksheet_name())

    def workbook_load_error_tests(self):

        workbook_obj = WorkbookWithHeader()
        self.assertRaises(WorkbookNotValid, workbook_obj.load_workbook, 'TestPath')

    def build_tab_list_by_type_tests(self):

        registered_tab_type_obj = RegisteredTabType()
        registered_tab_type_obj.registered_one_keyword_to_multiple_type(True)

        registered_tab_type_obj.registered_tab_type_identify('TestCases', 'TestFileName')
        registered_tab_type_obj.registered_tab_type_identify('Configs', 'ConfigName')

        workbook_obj = WorkbookWithHeader()
        workbook_obj.set_tab_type_list(registered_tab_type_obj)
        workbook_obj.set_default_header_row_no(0)
        workbook_obj.load_workbook(self.workbook_path)

        self.assertIn("KeysInHeader", workbook_obj.get_tab_list_by_type('TestCases'))
        self.assertIn("KeysInRows", workbook_obj.get_tab_list_by_type('TestCases'))
        self.assertIn("Configs", workbook_obj.get_tab_list_by_type('Configs'))

        registered_tab_type_obj.registered_tab_type_identify('TestCases', 'ConfigName')
        workbook_obj.build_tab_list_by_type()

        self.assertIn("KeysInHeader", workbook_obj.get_tab_list_by_type('TestCases'))
        self.assertIn("KeysInRows", workbook_obj.get_tab_list_by_type('TestCases'))
        self.assertIn("Configs", workbook_obj.get_tab_list_by_type('TestCases'))
        self.assertIn("Configs", workbook_obj.get_tab_list_by_type('Configs'))


def workbook_with_header_tests():
    suite = unittest.TestSuite()
    suite.addTest(WorkbookWithHeaderTest('workbook_load_tests'))
    suite.addTest(WorkbookWithHeaderTest('workbook_load_error_tests'))
    suite.addTest(WorkbookWithHeaderTest('build_tab_list_by_type_tests'))
    return suite


class RegisteredTabTypeTest(unittest.TestCase):

    def setUp(self):
        print('Setup for RegisteredTabTypeTest')

    def tearDown(self):
        print('Teardown for RegisteredTabTypeTest')

    def registered_one_keyword_to_multiple_type_test(self):
        registered_tab_type_obj = RegisteredTabType()
        self.assertFalse(registered_tab_type_obj.multiple_type_allowed)

        registered_tab_type_obj.registered_one_keyword_to_multiple_type(True)
        self.assertTrue(registered_tab_type_obj.multiple_type_allowed)

        registered_tab_type_obj.registered_one_keyword_to_multiple_type('')
        self.assertTrue(registered_tab_type_obj.multiple_type_allowed)

        registered_tab_type_obj.registered_one_keyword_to_multiple_type(' ')
        self.assertFalse(registered_tab_type_obj.multiple_type_allowed)

    def registered_tab_type_identify_tests(self):

        registered_tab_type_obj = RegisteredTabType()

        registered_tab_type_obj.registered_tab_type_identify('Type1', 'Keyword for Type1')
        self.assertIn('Type1', registered_tab_type_obj.type_identify_list.keys())
        self.assertIn('Keyword for Type1', registered_tab_type_obj.type_identify_list['Type1'])
        self.assertIn('Type1', registered_tab_type_obj.get_registered_tab_type_list())
        self.assertIn('Keyword for Type1', registered_tab_type_obj.get_identifying_keywords_by_tab_type('Type1'))

        registered_tab_type_obj.registered_tab_type_identify('Type2', 'Keyword for Type2')
        self.assertIn('Type2', registered_tab_type_obj.type_identify_list.keys())
        self.assertIn('Keyword for Type2', registered_tab_type_obj.type_identify_list['Type2'])
        self.assertIn('Type2', registered_tab_type_obj.get_registered_tab_type_list())
        self.assertIn('Keyword for Type2', registered_tab_type_obj.get_identifying_keywords_by_tab_type('Type2'))

        registered_tab_type_obj.registered_tab_type_identify('Type1', 'Keyword2 for Type1')
        self.assertIn('Type1', registered_tab_type_obj.type_identify_list.keys())
        self.assertIn('Keyword2 for Type1', registered_tab_type_obj.type_identify_list['Type1'])
        self.assertIn('Type1', registered_tab_type_obj.get_registered_tab_type_list())
        self.assertIn('Keyword2 for Type1', registered_tab_type_obj.get_identifying_keywords_by_tab_type('Type1'))

        registered_tab_type_obj.registered_tab_type_identify('Type2', 'Keyword for Type1')
        self.assertNotIn('Keyword for Type1', registered_tab_type_obj.type_identify_list['Type2'])

        registered_tab_type_obj.registered_one_keyword_to_multiple_type(True)
        registered_tab_type_obj.registered_tab_type_identify('Type2', 'Keyword for Type1')
        self.assertIn('Keyword for Type1', registered_tab_type_obj.type_identify_list['Type2'])
        self.assertIn('Keyword for Type1', registered_tab_type_obj.get_identifying_keywords_by_tab_type('Type2'))
        self.assertIn('Keyword for Type2', registered_tab_type_obj.get_identifying_keywords_by_tab_type('Type2'))

    def registered_tab_type_identify_error_tests(self):

        registered_tab_type_obj = RegisteredTabType()
        registered_tab_type_obj.registered_tab_type_identify('', '')
        self.assertEqual(0, len(registered_tab_type_obj.type_identify_list.keys()))

        registered_tab_type_obj.registered_tab_type_identify('TabType1', '')
        self.assertEqual(0, len(registered_tab_type_obj.type_identify_list.keys()))

    def unregistered_tab_type_identify_tests(self):

        registered_tab_type_obj = RegisteredTabType()

        registered_tab_type_obj.registered_tab_type_identify('Type1', 'Keyword for Type1')
        registered_tab_type_obj.registered_tab_type_identify('Type2', 'Keyword for Type2')
        self.assertIn('Type2', registered_tab_type_obj.type_identify_list.keys())
        self.assertIn('Keyword for Type2', registered_tab_type_obj.type_identify_list['Type2'])

        registered_tab_type_obj.unregistered_tab_type_identify('Type2', 'Keyword for Type2')
        self.assertIn('Type2', registered_tab_type_obj.type_identify_list.keys())
        self.assertNotIn('Keyword for Type2', registered_tab_type_obj.type_identify_list['Type2'])

        registered_tab_type_obj.unregistered_tab_type_identify('Type1', 'Keyword for Type1')
        self.assertIn('Type1', registered_tab_type_obj.type_identify_list.keys())
        self.assertNotIn('Keyword for Type1', registered_tab_type_obj.type_identify_list['Type1'])

        registered_tab_type_obj.unregistered_tab_type_identify('Type2', 'Keyword for Type2')
        registered_tab_type_obj.unregistered_tab_type_identify('Type1', 'Keyword for Type1')

    def get_tab_type_by_keywords_tests(self):

        registered_tab_type_obj = RegisteredTabType()
        registered_tab_type_obj.registered_one_keyword_to_multiple_type(True)

        registered_tab_type_obj.registered_tab_type_identify('Type1', 'Keyword for Type1')
        registered_tab_type_obj.registered_tab_type_identify('Type2', 'Keyword for Type2')
        registered_tab_type_obj.registered_tab_type_identify('Type2', 'Keyword for Type1')

        self.assertIn('Type1', registered_tab_type_obj.get_tab_type_by_keywords(['Keyword for Type1']))
        self.assertIn('Type2', registered_tab_type_obj.get_tab_type_by_keywords(['Keyword for Type1']))
        self.assertIn('Type2', registered_tab_type_obj.get_tab_type_by_keywords(['Keyword for Type2']))
        self.assertIn('Type2', registered_tab_type_obj.get_tab_type_by_keywords(['Keyword for Type1','Keyword for Type2']))
        self.assertIn('Type2', registered_tab_type_obj.get_tab_type_by_keywords(['Keyword for Type2','Keyword for Type3']))
        self.assertEqual([], registered_tab_type_obj.get_tab_type_by_keywords(['Keyword for Type3','Keyword for Type4']))
        self.assertEqual([], registered_tab_type_obj.get_tab_type_by_keywords([]))


def registered_tab_type_tests():
    suite = unittest.TestSuite()
    suite.addTest(RegisteredTabTypeTest('registered_one_keyword_to_multiple_type_test'))
    suite.addTest(RegisteredTabTypeTest('registered_tab_type_identify_tests'))
    suite.addTest(RegisteredTabTypeTest('registered_tab_type_identify_error_tests'))
    suite.addTest(RegisteredTabTypeTest('unregistered_tab_type_identify_tests'))
    suite.addTest(RegisteredTabTypeTest('get_tab_type_by_keywords_tests'))
    return suite


class WorksheetWithHeaderTest(unittest.TestCase):

    workbook_path = os.path.abspath(os.path.join(os.path.dirname(os.getcwd()), 'samp', 'Sample_Spread_with_Header.xlsx'))
    workbook_obj = xlrd.open_workbook(workbook_path)

    def setUp(self):
        print('Setup for WorksheetWithHeaderTest')

    def tearDown(self):
        print('Teardown for WorksheetWithHeaderTest')

    def load_worksheet_test(self):

        worksheet_obj = WorksheetWithHeader()
        worksheet_obj.set_header_row_number()
        worksheet_obj.load_worksheet(self.workbook_obj.sheet_by_name("KeysInHeader"))

        self.assertIsInstance(worksheet_obj, WorksheetWithHeader)
        self.assertEqual('KeysInHeader', worksheet_obj.get_worksheet_name())
        self.assertEqual(15, len(worksheet_obj.get_keyword_list()))
        self.assertIn("TestArea", worksheet_obj.get_keyword_list())
        self.assertIn("Parameters", worksheet_obj.get_keyword_list())
        self.assertIn("Key1", worksheet_obj.get_keyword_list())


def worksheet_with_header_tests():
    suite = unittest.TestSuite()
    suite.addTest(WorksheetWithHeaderTest('load_worksheet_test'))
    return suite


if __name__ == '__main__':
        runner = unittest.TextTestRunner()

        runner.run(workbook_with_header_tests())

        runner.run(registered_tab_type_tests())

        runner.run(worksheet_with_header_tests())
