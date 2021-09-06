## Uppraisal

This Python package uses the Canvas LMS API to upload Canvas assignment grades and comments to a Canvas course.

### Installation

You can install the released version of Uppraisal from GitHub with:

```shell 
$ pip install https://github.com/cisglee/uppraisal/zipball/master
```

Afterwards, go to the package folder and install the dependencies:

```shell
$ pip install -r requirements.txt
```

### Example

To use the package, proceed as follows:

1. Create a workbook with a `user_id` column containing the Canvas IDs of the students, a column containing comments (
   default name `submission_comment`), and a column containing grades (default name `grade`). Next, get an access token
   from your Canvas profile settings (for more information, see
   https://canvas.instructure.com/courses/785215/pages/getting-started-with-the-api).
2. Get the relevant course id from Canvas (can be found in the URL when you go to the course in your browser).
3. Get the relevant assignment id from Canvas (can be found in the URL when you go to the assignment in your browser).
4. Finally, run the package as follows, where `at` denotes access token, the `filepath` points to the Excel workbook
   with the assignment data, the `course_id` is the ID for the respective course and `assignment_id` is the ID for the
   respective assignment:

    ```python
    from uppraisal.uppraisal import upload_appraisals
    
    canvas_access_token = "VOIC57KB4OP35PXHTR1BI152F9XMF7683IAQG5SBRFVZBRUFHJIYPBEYTKI9J6LH69UFM3"
    filepath = "./assignment_data.xlsx"
    course_id = 101010
    assignment_id = 101010
    result = upload_appraisals(canvas_access_token, filepath, course_id, assignment_id)
    print(result)
    ```
