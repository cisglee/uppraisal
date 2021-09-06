#!/usr/bin/env python
# -*- coding: utf-8 -*-
import warnings
from copy import deepcopy

import requests
import xlrd
import xlwt


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
    :type has_headers: bool
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
    if worksheet is None:
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

            if has_headers and return_headers:
                data.append(org_fieldnames)

        else:
            row_data = deepcopy(data_template)
            curr_cell = -1
            while curr_cell < num_cells:
                curr_cell += 1
                fieldnr = fieldmap[curr_cell]
                if fieldnr is not None:
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


def putexcel(data, out_filepath, worksheet_names=None):
    """Method for writing data from a list of lists to Excel

    :Examples:

    >>> dat = [["Header 1", "Header 2", "Header 3"], ["row 1 cell 1", "row 1 cell 2", "row 1 cell 3"], ["row 2 cell 1", "row 2 cell 2", "row 3 cell 3"], ["row 3 cell 1", "row 3 cell 2", "row 3 cell 3"]]
    >>> putexcel(dat, "./example1.xls") # Example writing to a single worksheet
    >>> dat = [[["Header 1", "Header 2", "Header 3"], ["row 1 cell 1", "row 1 cell 2", "row 1 cell 3"], ["row 2 cell 1", "row 2 cell 2", "row 3 cell 3"], ["row 3 cell 1", "row 3 cell 2", "row 3 cell 3"]], [["Header 4", "Header 5", "Header 6"], ["row 4 cell 1", "row 4 cell 2", "row 4 cell 3"], ["row 5 cell 1", "row 5 cell 2", "row 5 cell 3"], ["row 6 cell 1", "row 6 cell 2", "row 6 cell 3"]], [["Header 7", "Header 8", "Header 9"], ["row 7 cell 1", "row 7 cell 2", "row 7 cell 3"], ["row 8 cell 1", "row 8 cell 2", "row 8 cell 3"], ["row 9 cell 1", "row 9 cell 2", "row 9 cell 3"]]]
    >>> putexcel(dat, "./example2.xls", worksheet_names=["WS1", "WS2", "WS3"]) # Example writing to multiple worksheets

    :param data: List of lists containing the data that should be written to Excel.
    :type data: list
    :param out_filepath: File path pointing to location and file name to which the data should be written.
    :type out_filepath: str
    :param worksheet_names: List of names of worksheets to which the data should be written or string name for the
    worksheet.
    :type worksheet_names: list or str
    :return: None
    """
    wb = xlwt.Workbook()
    l = 3 if isinstance(data[0][0], list) else 2
    if l == 2:
        # If data is list of rows, structure as if it were a list containing lists of rows
        if worksheet_names:
            if isinstance(worksheet_names, str):
                worksheet_names = [worksheet_names]
        data = [data]
        l = 3

    if l == 3:
        # If data is list containing lists of rows create a worksheet for every list of rows
        # Check if there are names for every worksheet
        if worksheet_names:
            if len(data) != len(worksheet_names):
                raise IOError("Data structure indicates {} worksheet, but {} worksheet name{} been provided".format(
                    len(data), len(worksheet_names), "s have" if len(worksheet_names) > 1 else " has"
                ))
        else:
            worksheet_names = ["Sheet{}".format(i + 1) for i in range(len(data))]

        for wsname, wsdat in zip(worksheet_names, data):
            ws = wb.add_sheet(wsname)
            for row_j, row in enumerate(wsdat):
                for cell_i, cell in enumerate(row):
                    ws.write(row_j, cell_i, cell)

        wb.save(out_filepath)

    return


def list_submissions(cv_access_token, cv_id_course, cv_id_assignment, sort_by='user_sortable_name',
                     to_excel='./assignment_data.xls', select_columns=(
                'user_id', 'user_sortable_name', 'grade',
                'score', 'submitted_at', 'preview_url', 'attachments'
        )):
    """Retrieve submission meta data for a Canvas assignment

    :param cv_access_token: Canvas access token (generally 70 characters in length; see
    https://canvas.instructure.com/courses/785215/pages/getting-started-with-the-api).
    :type cv_access_token: str
    :param cv_id_course: Course identifier (see Canvas course URL).
    :type cv_id_course: int
    :param cv_id_assignment: Assignment identifier (see Canvas assignment URL).
    :type cv_id_assignment: int
    :param sort_by: Assignment meta data column by which the return data should be sorted (e.g., 'submitted_at'). Default
    is 'user_sortable_name'.
    :type sort_by: str
    :param select_columns: List with column names to return or None if all meta data should be returned.
    :type select_columns: None or list
    :return: List of dictionary items containing meta data for assignments submitted to assignment id `cv_id_assignment`
    :rtype: list
    """

    def _filter_assignments(rbuffer):
        assignments = []
        raw = rbuffer.json()
        for assignment in raw:
            if assignment['submitted_at']:
                if 'user' in assignment:
                    for key in assignment['user'].keys():
                        assignment['user_' + key] = assignment['user'][key]
                    del assignment['user']
                if 'attachments' in assignment:
                    attachments = []
                    for attachment in assignment['attachments']:
                        attachments.append(attachment['filename'])
                    assignment['attachments'] = '; '.join(attachments)
                if select_columns:
                    remove_columns = []
                    for col in assignment:
                        if col not in select_columns:
                            remove_columns.append(col)
                    for col in remove_columns:
                        del assignment[col]
                assignments.append(assignment)
        return assignments

    if sort_by and select_columns and sort_by not in select_columns:
        raise IOError('The column specified for sortby must be in the list of columns you would like to be returned')

    params = {'grouped': True, 'per_page': 100, 'include': 'user'}
    headers = {'Authorization': 'Bearer {}'.format(cv_access_token), 'Content-type': 'application/json'}
    assignments = []
    rbuffer = requests.get('https://canvas.eur.nl/api/v1/courses/{cid}/assignments/{aid}/submissions'.format(
        cid=cv_id_course,
        aid=cv_id_assignment
    ),
        headers=headers,
        params=params,
    )

    assignments += _filter_assignments(rbuffer)
    while rbuffer.links['current']['url'] != rbuffer.links['last']['url']:
        rbuffer = requests.get(
            rbuffer.links['next']['url'],
            headers=headers,
            params=params,
        )
        assignments += _filter_assignments(rbuffer)

    if sort_by in assignments[0]:
        assignments.sort(key=lambda k: k[sort_by])
    else:
        if type(sort_by) == str:
            warnings.warn('The specified sort key (i.e., {}) was not found'.format(sort_by))
        else:
            warnings.warn('The parameter sortby should be a string'.format(sort_by))

    headers = []
    if select_columns:
        headers = list(select_columns)
    else:
        for assignment in assignments:
            headers.append(assignment.keys())
        headers = list(set(headers))

    assignment_list = [headers]
    for assignment in assignments:
        assignment_row = []
        for header in headers:
            assignment_row.append(assignment[header])
        assignment_list.append(assignment_row)

    if to_excel:
        putexcel(assignment_list, to_excel)

    return assignments


def chunker(seq, size):
    """Chunks a list of items to help reduce the API load

    :source: https://stackoverflow.com/a/434328
    :param seq: array over which to loop
    :rtype seq: list
    :param size: size of chunks
    :rtype size: int
    :return: chunked list
    :rtype: list
    """
    # (in python 2 use xrange() instead of range() to avoid allocating a list)
    return (seq[pos:pos + size] for pos in range(0, len(seq), size))
