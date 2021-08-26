#!/usr/bin/env python
# -*- coding: utf-8 -*-
from copy import deepcopy

import xlrd


def getexcel(in_filepath, worksheet, include_fieldnames=None, exclude_fieldnames=None, has_headers=True,
             return_headers=True, capital_sensitive=False, ignore_missing_headers=False):
    """Method for reading data from Excel files to a list of lists.

    :param in_filepath: A filepath string referring to an Excel file.
    :type in_filepath: str
    :param worksheet: The name or index number of the worksheet from which data should be retrieved.
    :type worksheet: str or int
    :param include_fieldnames: List of fieldnames that should be returned from Excel workbook
    :type include_fieldnames: list
    :param exclude_fieldnames: list of fieldnames that should be ignored when retrieving data from Excel workbook
    :type exclude_fieldnames: list
    :param has_headers: Boolean indicating whether workbook has headers in the first row (True) or not (False)
    :type has_heasers: bool
    :param return_headers: Boolean indicating whether the headers should be returned
    :type return_headers: bool
    :param capital_sensitive: Boolean indicating whether include_fieldnames or exclude_fieldnames should be sensitive to capitalization
    :type capital_sensitive: bool
    :param ignore_missing_headers: Boolean indicating whether missing headers from include_fieldnames or exclude_fieldnames should be ignored (True) or not (False)
    :type ignore_missing_headers: bool
    :return: A list of lists containing the data from the worksheet.
    :rtype: list
    """
    # TODO: Check fieldnames
    # TODO: Check for duplicate headers
    if (include_fieldnames or exclude_fieldnames) and not has_headers:
        raise Exception("Cannot include or exclude data from Excel file by column names when there are no headers.")
    if include_fieldnames and exclude_fieldnames:
        raise Exception("Cannot import data based on inclusion and exclusion simultaneously.")
    if return_headers and not has_headers:
        raise Exception("Cannot return headers from Excel file when there are no headers.")
    if capital_sensitive or ignore_missing_headers and not has_headers:
        raise Warning(
            "Arguments \"capital_sensitive\" and \"ignore_missing_headers\" have no effect when there are no headers.")

    workbook = xlrd.open_workbook(in_filepath)
    if not worksheet:
        raise IOError("The name or index number of the worksheet has not been specified.")
    elif type(worksheet) == int:
        worksheet = workbook.sheet_by_index(worksheet)
    elif type(worksheet) == str:
        worksheet = workbook.sheet_by_name(worksheet)

    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    curr_row = -1

    if include_fieldnames:
        org_fieldnames = include_fieldnames
        include_fieldnames = [field.lower().strip() if capital_sensitive is False else field.strip() for field in
                              include_fieldnames]
        data_template = [None for field in include_fieldnames]
        fieldmap = [None for i in range(0, num_cells + 1)]
    elif exclude_fieldnames:
        org_fieldnames = include_fieldnames
        exclude_fieldnames = [field.lower().strip() if capital_sensitive is False else field.strip() for field in
                              exclude_fieldnames]
        data_template = [None for field in range(0, num_cells + 1 - len(exclude_fieldnames))]
        fieldmap = [None for i in range(0, num_cells + 1)]
    else:
        data_template = [None for field in range(0, num_cells + 1)]
        fieldmap = [i for i in range(0, num_cells + 1)]

    data = []
    fieldnames_found = []
    headers = []
    while curr_row < num_rows:
        curr_row += 1

        if curr_row == 0 and has_headers:
            if include_fieldnames:
                curr_cell = -1
                while curr_cell < num_cells:
                    curr_cell += 1
                    cell_value = worksheet.cell_value(curr_row,
                                                      curr_cell).lower().strip() if capital_sensitive is False else worksheet.cell_value(
                        curr_row, curr_cell).strip()
                    if cell_value in include_fieldnames:
                        fieldmap[curr_cell] = include_fieldnames.index(cell_value)
                        fieldnames_found.append(cell_value)
            elif exclude_fieldnames:
                include_fieldnames = []
                curr_cell = -1
                while curr_cell < num_cells:
                    curr_cell += 1
                    cell_value = worksheet.cell_value(curr_row,
                                                      curr_cell).lower().strip() if capital_sensitive is False else worksheet.cell_value(
                        curr_row, curr_cell).strip()
                    if cell_value not in exclude_fieldnames:
                        include_fieldnames.append(cell_value)
                        fieldnr = max(fieldmap) + 1 if 0 in fieldmap else 0
                        fieldmap[curr_cell] = fieldnr
                        fieldnames_found.append(cell_value)
            else:
                # If not include or exclude fieldnames, include_fieldnames is set to all fieldnames
                org_fieldnames = worksheet.row_values(curr_row)

            if has_headers == True and return_headers == True:
                data.append(org_fieldnames)

        else:
            row_data = deepcopy(data_template)
            curr_cell = -1
            while curr_cell < num_cells:
                curr_cell += 1
                fieldnr = fieldmap[curr_cell]
                if fieldnr != None:
                    row_data[fieldnr] = worksheet.cell_value(curr_row, curr_cell)

            data.append(row_data)

    if ignore_missing_headers is False:
        missing = None
        if include_fieldnames:
            missing = set(include_fieldnames) - set(fieldnames_found)
        elif exclude_fieldnames:
            missing = set(exclude_fieldnames) - set(fieldnames_found)
        if missing:
            raise IOError("The following fieldnames could not be found: {}".format(", ".join(missing)))

    return data
