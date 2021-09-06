#!/usr/bin/env python
# -*- coding: utf-8 -*-
import json
import time

import requests
from tqdm import tqdm

from uppraisal.helpers import chunker
from uppraisal.helpers import getexcel


def upload_appraisal(cv_access_token, workbook_path, cv_id_course, cv_id_assignment,
                     submission_comment_header="submission_comment", grade_header="grade", workbook_tab='Sheet1'):
    """Upload assignment comments and grades from an excel workbook, one assignment at a time.

    :param cv_access_token: Canvas access token (generally 70 characters in length; see
    https://canvas.instructure.com/courses/785215/pages/getting-started-with-the-api).
    :type cv_access_token: str
    :param workbook_path: File path to workbook containing grade and comment data, with Canvas user_id's included
    under the header "user_id".
    :type workbook_path: str
    :param cv_id_course: Course identifier (see Canvas course URL).
    :type cv_id_course: int
    :param cv_id_assignment: Assignment identifier (see Canvas assignment URL).
    :type cv_id_assignment: int
    :param submission_comment_header: Header of the submission comment column in the Excel workbook. Default is
    "submission_comment".
    :type submission_comment_header: str
    :param grade_header: Header of the grades column in the Excel workbook. Default is "grade".
    :type grade_header: str
    :param workbook_tab: The name of the worksheet from which data should be retrieved.
    :type workbook_tab: str
    :return: Dictionary with updated assignment meta data or an error message
    :rtype: dict
    """
    fields = ["user_id", grade_header, submission_comment_header]
    data = getexcel(workbook_path, workbook_tab, include_fieldnames=fields, return_headers=False)
    for i, row in enumerate(data):
        data[i][0] = int(row[0])
    r_data = {}
    for cv_id_user, grade, submission_comment in data:
        params = {
            "comment": {"text_comment": submission_comment},
            "submission": {"posted_grade": grade}
        }
        r = requests.put('https://canvas.eur.nl/api/v1/courses/{cid}/assignments/{aid}/submissions/{uid}'.format(
            cid=cv_id_course,
            aid=cv_id_assignment,
            uid=cv_id_user
        ),
            headers={'Authorization': 'Bearer ' + cv_access_token, 'Content-type': 'application/json'},
            json=params,
        )
        time.sleep(.02)  # TODO: Check whether the delay is sufficient
        r_data[cv_id_user] = json.loads(r.content.decode('utf-8'))
    return r_data


def upload_appraisals(cv_access_token, workbook_path, cv_id_course, cv_id_assignment,
                      submission_comment_header="submission_comment", grade_header="grade", workbook_tab='Sheet1'):
    """Upload assignment comments and grades from an excel workbook, with multiple assignments per post.

    :param cv_access_token: Canvas access token (generally 70 characters in length; see
    https://canvas.instructure.com/courses/785215/pages/getting-started-with-the-api).
    :type cv_access_token: str
    :param workbook_path: File path to workbook containing grade and comment data, with Canvas user_id's
    included under the header "user_id".
    :type workbook_path: str
    :param cv_id_course: Course identifier (see Canvas course URL).
    :type cv_id_course: int
    :param cv_id_assignment: Assignment identifier (see Canvas assignment URL).
    :type cv_id_assignment: int
    :param submission_comment_header: Header of the submission comment column in the Excel workbook. Default is
    "submission_comment".
    :type submission_comment_header: str
    :param grade_header: Header of the grades column in the Excel workbook. Default is "grade".
    :type grade_header: str
    :param workbook_tab: The name of the worksheet from which data should be retrieved.
    :type workbook_tab: str
    :return: Dictionary with updated assignment meta data or an error message
    :rtype: dict
    """
    fields = ["user_id", grade_header, submission_comment_header]
    data = getexcel(workbook_path, workbook_tab, include_fieldnames=fields, return_headers=False)
    for i, row in enumerate(data):
        data[i][0] = int(row[0])
    r_data = []
    for data_chunk in tqdm(chunker(data, 100)):
        params = {"grade_data": {}}
        for cv_id_user, grade, submission_comment in data_chunk:
            params["grade_data"][int(cv_id_user)] = {
                # TODO: Create a test to verify robustness against floating point imprecision errors
                "posted_grade": str(grade),
                # TODO: Check encoding and non-ASCII character issues
                "text_comment": submission_comment
            }
        r = requests.post(
            'https://canvas.eur.nl/api/v1/courses/{cid}/assignments/{aid}/submissions/update_grades'.format(
                cid=cv_id_course,
                aid=cv_id_assignment
            ),
            headers={'Authorization': 'Bearer ' + cv_access_token, 'Content-type': 'application/json'},
            json=params,
        )
        r_dat = json.loads(r.content.decode('utf-8'))
        state = 'queued'
        while state == 'queued':
            time.sleep(30)
            progress = requests.get(r_dat['url'],
                                    headers={'Authorization': 'Bearer ' + cv_access_token,
                                             'Content-type': 'application/json'})
            progress = json.loads(progress.content.decode('utf-8'))
            state = progress['workflow_state']
        r_data.append(r_dat)
    return r_data
