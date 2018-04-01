# encoding: utf-8

import os
import xlrd
from internallogging import *


class WorkbookNotValid(Exception):

    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class WorkbookWithHeader(object):

    default_header_row_no = None            # Default settings of the row number of header locate
    type_identify_list = None               # List of the type and related identifying keywords.
    workbook_obj = None                     # Object of workbook used in this class
    workbook_path = None                    # Path of workbook located.
    tab_list = None                         # list of tab name in opened workbook.
    internal_logger = None                  # Internal logger object.

    def __init__(self, workbook_path=None):

        self.internal_logger = get_internal_logger()
        self.type_identify_list = dict()
        self.default_header_row_no = 0
        self.tab_list = []

        if workbook_path: self.load_workbook(workbook_path)

    def load_workbook(self, workbook_path):
        """
        Open specified workbook file and load all tab name to the list.

        :param workbook_path: Path of workbook located.
        :return: None. Will raise WorkbookNotValid exception if specified path is invalid or not a workbook file.
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
            self.tab_list = self.workbook_obj.sheet_names()

    def registered_tab_type_identify(self, tab_type, identifying_keywords, multiple_type_allowed=False):
        """
        A tab type identifying keywords is one of element in header row to indicate the tab type.
        e.g.:
            Define the element 'Parameter' in header row to indicate this tab is used to store test cases.
            Define the element 'Config Name' in header row to indicate this tab is used to store test config.

        By default one keyword can be registered for ONLY one type of tab to avoid confusion.
        This rule could be changed by set multiple_type_allowed to True

        The data structure of type_identify_list is like:
        {
            <type of tab> :
                        [
                            <list of identifying keywords for this type.>
                        ]
        }

        :param tab_type: A user defined tab type string.
        :param identifying_keywords: The keyword related to tab type. Should be a element in header row.
        :param multiple_type_allowed: Whether to allow one keyword to be registered for more than one type of tab.
        :return: None. The identifying keyword was add to related tab type in the type_identify_list
        """

        if not tab_type: return
        if not identifying_keywords: return

        # Check if identifying_keywords was already registered.
        if not multiple_type_allowed:
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
        pass