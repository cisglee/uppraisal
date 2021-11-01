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
4. Finally, run the package as follows, where the `filepath` points to the Excel workbook with the assignment data,
   the `-c` is the ID for the respective course, `t` denotes access token, and `-a` is the ID for the respective
   assignment:

    ```shell
    $ ./main.py -c [course_id] -a [assignment_id] -t [access_token] filepath
    ```

 
