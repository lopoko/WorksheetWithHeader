# encoding: utf-8

import os
import xlrd
from internallogging import *


class WorkbookNotValid(Exception):

    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class WorksheetNotFound(Exception):

    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class RegisteredTabType(object):
    """
    Registered tab type with keyword in header row.
    """

    type_identify_list = None               # List of the type and related identifying keywords.
    internal_logger = None                  # Internal logger object.
    multiple_type_allowed = None            # Whether to allow one keyword to be registered for more than one type of tab.

    def __init__(self):
        self.type_identify_list = dict()
        self.multiple_type_allowed = False

    def registered_one_keyword_to_multiple_type(self, multiple_type_allowed):
        if multiple_type_allowed: self.multiple_type_allowed = multiple_type_allowed is True

    def registered_tab_type_identify(self, tab_type, identifying_keywords):
        """
        A tab type identifying keywords is one of element in header row to indicate the tab type.
        e.g.:
            Define the element 'Parameter' in header row to indicate this tab is used to store test cases.
            Define the element 'Config Name' in header row to indicate this tab is used to store test config.

        By default one keyword can be registered for ONLY one type of tab to avoid confusion.
        This rule could be changed by set multiple_type_allowed to True

        The data structure of type_identify_list is like:
        {
            <type of tab>: [ <keywords for this type>, <keywords for this type>, ...... ],
            <type of tab>: [ <keywords for this type>, <keywords for this type>, ...... ],
            ......
        }

        :param tab_type: A user defined tab type string.
        :param identifying_keywords: The keyword related to tab type. Should be a element in header row.
        :return: None. The identifying keyword was add to related tab type in the type_identify_list
        """

        if not tab_type: return
        if not identifying_keywords: return

        # Check if identifying_keywords was already registered.
        if not self.multiple_type_allowed:
            for registered_tab_type in self.type_identify_list.keys():
                if identifying_keywords in self.type_identify_list[registered_tab_type]: return

        # Registered the keyword.
        if tab_type not in self.type_identify_list.keys(): self.type_identify_list[tab_type] = []
        if identifying_keywords not in self.type_identify_list[tab_type]:
            self.type_identify_list[tab_type].append(identifying_keywords)

    def unregistered_tab_type_identify(self, tab_type, identifying_keywords):
        """
        Unregistered the keyword from specified tab type.

        :param tab_type: A user defined tab type string.
        :param identifying_keywords: The keyword related to tab type, which need be unregistered.
        :return: None. specified keyword will be removed from type_identify_list
        """
        if (tab_type in self.type_identify_list.keys()) and (identifying_keywords in self.type_identify_list[tab_type]):
            self.type_identify_list[tab_type].remove(identifying_keywords)

    def get_identifying_keywords_by_tab_type(self, tab_type_name):
        """
        Return the registered identifying keywords by given tab type.

        :param tab_type_name: given tab type for querying
        :return: list of registered identifying keywords for given tab type.
        """
        if tab_type_name in self.type_identify_list.keys():
            return self.type_identify_list[tab_type_name].copy()

    def get_registered_tab_type_list(self):
        """
        Return the list of registered tab type.

        :return: the list of registered tab type.
        """
        return self.type_identify_list.keys()

    def get_tab_type_by_keywords(self, keyword_list):
        """
        Identify the tab type by given list of keyword.
        Multiple tab type will be returned if keywords in more than one kind of type was found.

        :param keyword_list: list of keyword need be parsed.
        :return: list of found tab type.
        """

        tab_type_list = []

        if isinstance(keyword_list, list):
            for tab_type in self.type_identify_list.keys():
                for tab_keyword in self.type_identify_list[tab_type]:
                    if tab_keyword in keyword_list:
                        tab_type_list.append(tab_type)
                        break

        return tab_type_list.copy()


class WorkbookWithHeader(object):
    """
    Object to handle workbook with header.
    """

    default_header_row_no = None            # Default settings of the row number of header locate
    registered_tab_type = None              # Object of registered tab type
    workbook_obj = None                     # Object of workbook used in this class
    workbook_path = None                    # Path of workbook located.
    tab_list_by_type = None                 # dict of tab name in opened workbook, grouped by tab type in keys.
    worksheet_list_by_name = None           # dict of WorksheetWithHeader objects relate to worksheets in the workbook.
    internal_logger = None                  # Internal logger object.

    def __init__(self):

        self.internal_logger = get_internal_logger()
        self.default_header_row_no = 0
        self.tab_list_by_type = dict()
        self.worksheet_list_by_name = dict()

    def set_tab_type_list(self, tab_type_list):
        """
        Give a RegisteredTabType to the workbook object.

        :param tab_type_list: A RegisteredTabType object need passed to workbook object.
        :return: None.
        """

        if isinstance(tab_type_list, RegisteredTabType):
            self.registered_tab_type = tab_type_list

    def set_default_header_row_no(self, header_row_no):
        """
        Set the default row number of header row in the worksheet.

        :param header_row_no: default row number of the header row in worksheet.
        :return: None. default_header_row_no was changed if header_row_no is valid (is positive integer)
        """

        try:
            if int(header_row_no) >= 0: self.default_header_row_no = header_row_no
        except Exception:
            self.internal_logger.exception("Error occurs when set default header row number to %s" % header_row_no)

    def load_workbook(self, workbook_path):
        """
        Open specified workbook file and load all tab name to the list.

        :param workbook_path: Path of workbook located.
        :return: None.
                    workbook_path and list_of_worksheet will be updated.
                    Will raise WorkbookNotValid exception if specified path is invalid or not a workbook file.
        """

        if workbook_path:
            if not os.path.exists(workbook_path):
                raise WorkbookNotValid(workbook_path, 'Workbook file Not Found.')
            else:
                try:
                    self.workbook_obj = xlrd.open_workbook(workbook_path)
                except Exception as workbook_error:
                    self.internal_logger.exception('Failed to open Workbook file %s' % workbook_path)
                    raise WorkbookNotValid(workbook_path, 'Failed to open Workbook file: %s' % workbook_error)

            self.workbook_path = workbook_path

            # Add related WorksheetWithHeader object to the worksheet_list_by_name
            self.worksheet_list_by_name.clear()
            for tab_name in self.workbook_obj.sheet_names():
                worksheet_obj = WorksheetWithHeader()
                worksheet_obj.set_header_row_number(self.default_header_row_no)
                worksheet_obj.load_worksheet(self.workbook_obj.sheet_by_name(tab_name))
                self.worksheet_list_by_name[tab_name] = worksheet_obj

            # Build tab list by type after load.
            self.build_tab_list_by_type()

    def build_tab_list_by_type(self):
        """
        Build tab list by type:
            - Find keywords in each tab/worksheet
            - Detect the tab type list by calling RegisteredTabType.get_tab_type_by_keywords()
            - Group tab name by the returned tab type list.

        Data structure of tab_list_by_type:
        {
            <tab type>: [ <tab name>, <tab name>, ......],
            <tab type>: [ <tab name>, <tab name>, ......],
            ......
        }
        :return: None.  tab_list_by_type will be updated.
        """

        if self.registered_tab_type is None: return

        self.tab_list_by_type.clear()
        for worksheet_obj in self.worksheet_list_by_name.values():
            for tab_type in self.registered_tab_type.get_tab_type_by_keywords(worksheet_obj.get_keyword_list()):
                if tab_type not in self.tab_list_by_type.keys(): self.tab_list_by_type[tab_type] = []
                self.tab_list_by_type[tab_type].append(worksheet_obj.get_worksheet_name())

    def get_worksheet_by_name(self, worksheet_name):
        """
        Return the WorksheetWithHeader object with specified sheet name.

        :param worksheet_name: name of worksheet need be returned
        :return:  WorksheetWithHeader object with specified sheet name.
                   Will raise WorksheetNotFound exception if specified worksheet name was not found in tab list.
        """

        if worksheet_name and worksheet_name in self.worksheet_list_by_name.keys():
            return self.worksheet_list_by_name[worksheet_name]
        else:
            raise WorksheetNotFound(worksheet_name, 'Worksheet Not Found in workbook %s.' % self.workbook_path)

    def get_tab_list(self):
        """
        Return tab name list of current workbook.

        :return: tab name list of current workbook
        """

        return self.workbook_obj.sheet_names() if self.workbook_obj else []

    def get_tab_list_by_type(self, tab_type):
        """
        Return the list of tab which header row has keyword match the registered tab type identifier.

        :param tab_type: Name of the tab type in registered tab type list.
        :return: list of the tab name match the specified tab type.
        """

        if tab_type in self.tab_list_by_type.keys():
            return self.tab_list_by_type[tab_type].copy()
        else:
            return []


class WorksheetWithHeader(object):
    """
    Object to handle worksheet with header
    """

    header_row_no = None                # The row number of header locates in the worksheet
    keywords_in_header = None           # The list of keywords present in header row
    max_row_usage = None                # Use self.worksheet_object.nrows
    worksheet_object = None             # Related worksheet object.

    def __init__(self):

        self.header_row_no = 0
        self.keywords_in_header = []

    def set_header_row_number(self, header_row_no=0):
        """
        Set the row number where header located.

        :param header_row_no: row number of header located.
        :return: None, header_row_no will be updated.
        """

        if header_row_no >=0: self.header_row_no = header_row_no

    def load_worksheet(self, worksheet_obj):
        """
        Load worksheet then detect max row usage and keywords in header.

        :param worksheet_obj: Object of worksheet need be handled.
        :return: None, worksheet_object, and keywords_in_header will be handled.
        """

        if isinstance(worksheet_obj, xlrd.sheet.Sheet):
            self.worksheet_object = worksheet_obj
            self.keywords_in_header.clear()

            # load keywords in header
            for current_cell in self.worksheet_object.row(self.header_row_no):
                if current_cell.ctype not in [xlrd.XL_CELL_BLANK, xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_ERROR]:
                    self.keywords_in_header.append(current_cell.value)

    def get_keyword_list(self):
        """
        Return the keyword list load from the spreadsheet.

        :return: list of keyword load from the spreadsheet.
        """

        return self.keywords_in_header.copy()

    def get_worksheet_name(self):
        """
        Return the worksheet name
        :return: worksheet name from worksheet object.
        """

        return self.worksheet_object.name
