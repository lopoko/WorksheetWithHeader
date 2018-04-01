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

        self.assertIn('KeysInHeader', workbook_obj.tab_list)
        self.assertIn('KeysInRows', workbook_obj.tab_list)
        self.assertIn('Configs', workbook_obj.tab_list)

    def workbook_load_error_tests(self):

        self.assertRaises(WorkbookNotValid, WorkbookWithHeader, 'TestPath')

    def registered_tab_type_identify_tests(self):

        workbook_obj = WorkbookWithHeader()

        workbook_obj.registered_tab_type_identify('Type1','Keyword for Type1')
        self.assertIn('Type1', workbook_obj.type_identify_list.keys())
        self.assertIn('Keyword for Type1', workbook_obj.type_identify_list['Type1'])

        workbook_obj.registered_tab_type_identify('Type2','Keyword for Type2')
        self.assertIn('Type2', workbook_obj.type_identify_list.keys())
        self.assertIn('Keyword for Type2', workbook_obj.type_identify_list['Type2'])

        workbook_obj.registered_tab_type_identify('Type1','Keyword2 for Type1')
        self.assertIn('Type1', workbook_obj.type_identify_list.keys())
        self.assertIn('Keyword2 for Type1', workbook_obj.type_identify_list['Type1'])

        workbook_obj.registered_tab_type_identify('Type2','Keyword for Type1')
        self.assertNotIn('Keyword for Type1', workbook_obj.type_identify_list['Type2'])

        workbook_obj.registered_tab_type_identify('Type2','Keyword for Type1', True)
        self.assertIn('Keyword for Type1', workbook_obj.type_identify_list['Type2'])

    def registered_tab_type_identify_error_tests(self):

        workbook_obj = WorkbookWithHeader()
        workbook_obj.registered_tab_type_identify('', '')
        self.assertEqual(0, len(workbook_obj.type_identify_list.keys()))

        workbook_obj.registered_tab_type_identify('TabType1', '')
        self.assertEqual(0, len(workbook_obj.type_identify_list.keys()))


def workbook_with_header_tests():
    suite = unittest.TestSuite()
    suite.addTest(WorkbookWithHeaderTest('workbook_load_tests'))
    suite.addTest(WorkbookWithHeaderTest('workbook_load_error_tests'))
    suite.addTest(WorkbookWithHeaderTest('registered_tab_type_identify_tests'))
    suite.addTest(WorkbookWithHeaderTest('registered_tab_type_identify_error_tests'))
    return suite


if __name__ == '__main__':
        runner = unittest.TextTestRunner()
        runner.run(workbook_with_header_tests())
