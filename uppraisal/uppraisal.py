# -*- coding: utf-8 -*-

import json
import time
import typing as tp

import requests
from bs4 import BeautifulSoup
from tqdm import tqdm

from .constants import *
from .helpers import chunker
from .helpers import get_excel


def upload_appraisals(
        in_filepath: str,
        cv_access_token: str,
        cv_course_id: int,
        cv_assignment_id: int,
        html_format: bool = DEFAULT_HTML_FORMAT,
        submission_comment_header: str = DEFAULT_SUBMISSION_COMMENT_HEADER,
        grade_header: str = DEFAULT_GRADE_HEADER,
        workbook_tab: str = DEFAULT_WORKBOOK_TAB
) -> tp.List[tp.Any]:
    """
    Upload assignment comments and grades from an excel workbook, with multiple assignments per post.

    :param in_filepath: File path to workbook containing grade and comment data, with Canvas user_id's
    included under the header "user_id".
    :type in_filepath: str
    :param cv_access_token: Canvas access token (generally 70 characters in length; see
    https://canvas.instructure.com/courses/785215/pages/getting-started-with-the-api).
    :type cv_access_token: str
    :param cv_course_id: Course identifier (see Canvas course URL).
    :type cv_course_id: int
    :param cv_assignment_id: Assignment identifier (see Canvas assignment URL).
    :type cv_assignment_id: int
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

    fields = ['user_id', grade_header, submission_comment_header]
    data = get_excel(
        in_filepath,
        workbook_tab,
        include_fieldnames=fields,
        return_headers=False
    )

    for i, row in enumerate(data):
        data[i][0] = int(row[0])

    responses = []
    for data_chunk in tqdm(chunker(data)):
        params = {'grade_data': {}}
        for cv_id_user, grade, submission_comment in data_chunk:
            params['grade_data'][int(cv_id_user)] = {
                # TODO: Create a test to verify robustness against floating point imprecision errors
                'posted_grade': str(grade),
                # TODO: Check encoding and non-ASCII character issues
                'text_comment': submission_comment if html_format is False else BeautifulSoup(
                    submission_comment.replace("<br>", "\n"), features="html.parser").get_text()
            }
        url = CANVAS_UPDATE_GRADES_URL.format(cid=cv_course_id, aid=cv_assignment_id)
        response = requests.post(
            url,
            headers={
                'Authorization': 'Bearer ' + cv_access_token,
                'Content-type': 'application/json'
            },
            json=params
        )
        if response.status_code // 100 != 2:
            raise Exception(f'Got an error {response.status_code} from {url}: {response.text}')

        body = json.loads(response.text)
        state = 'queued'
        while state == 'queued':
            time.sleep(SLEEP_UPLOAD_APPRAISALS)
            url = body['url']
            response = requests.get(
                url,
                headers={
                    'Authorization': 'Bearer ' + cv_access_token,
                    'Content-type': 'application/json'
                }
            )
            if response.status_code // 100 != 2:
                raise Exception(
                    f'Got an error {response.status_code} from {url}: {response.text}')

            progress = json.loads(response.text)
            state = progress['workflow_state']

        responses.append(body)

    return responses


def upload_appraisal(
        cv_access_token: str,
        workbook_path: str,
        cv_course_id: int,
        cv_assignment_id: int,
        html_format: bool = DEFAULT_HTML_FORMAT,
        submission_comment_header: str = DEFAULT_SUBMISSION_COMMENT_HEADER,
        grade_header: str = DEFAULT_GRADE_HEADER,
        workbook_tab: str = DEFAULT_WORKBOOK_TAB
) -> tp.Dict[int, tp.Any]:
    """
    Upload assignment comments and grades from an excel workbook, one assignment at a time.

    :param cv_access_token: Canvas access token (generally 70 characters in length; see
    https://canvas.instructure.com/courses/785215/pages/getting-started-with-the-api).
    :type cv_access_token: str
    :param workbook_path: File path to workbook containing grade and comment data, with Canvas user_id's included
    under the header "user_id".
    :type workbook_path: str
    :param cv_course_id: Course identifier (see Canvas course URL).
    :type cv_course_id: int
    :param cv_assignment_id: Assignment identifier (see Canvas assignment URL).
    :type cv_assignment_id: int
    :param html_format: Indicate that submission comment is in HTML
    :type html_format: bool
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
    data = get_excel(workbook_path, workbook_tab, include_fieldnames=fields, return_headers=False)
    for i, row in enumerate(data):
        data[i][0] = int(row[0])
    responses = {}
    for cv_user_id, grade, submission_comment in data:
        params = {
            'comment': {'text_comment': submission_comment if html_format is False else BeautifulSoup(
                submission_comment.replace("<br>", "\n"), features="html.parser").get_text()},
            'submission': {'posted_grade': grade}
        }

        url = CANVAS_SUBMISSION_URL.format(
            cid=cv_course_id,
            aid=cv_assignment_id,
            uid=cv_user_id
        )
        response = requests.put(
            url,
            headers={
                'Authorization': 'Bearer ' + cv_access_token,
                'Content-type': 'application/json'
            },
            json=params,
        )
        if response.status_code // 100 != 2:
            raise Exception(f'Got an error {response.status_code} from {url}: {response.text}')

        time.sleep(SLEEP_UPLOAD_APPRAISAL)
        responses[cv_user_id] = json.loads(response.text)

    return responses
