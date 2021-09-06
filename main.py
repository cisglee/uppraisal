#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse

from uppraisal.uppraisal import upload_appraisals

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-t', '--token',
        help="Specify the Canvas access token",
        type=str,
        required=True
    )
    parser.add_argument(
        '-c', '--course',
        help="Specify the Canvas course ID",
        type=int,
        required=True
    )
    parser.add_argument(
        '-a', '--assignment',
        help="Specify the Canvas course ID",
        type=int,
        required=True
    )
    parser.add_argument(
        'filepath',
        help="Specify an Excel file with the results",
        type=str,
    )

    arguments = parser.parse_args()

    result = upload_appraisals(
        arguments.filepath,
        cv_access_token=arguments.token,
        cv_course_id=arguments.course,
        cv_assignment_id=arguments.assignment
    )

    print(result)
