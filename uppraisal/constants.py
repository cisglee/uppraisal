import typing as _tp

CANVAS_BASE_URL: str = 'https://canvas.eur.nl/api/v1'
CANVAS_COURSES_URL: str = CANVAS_BASE_URL + '/courses'
CANVAS_COURSE_URL: str = CANVAS_COURSES_URL + '/{cid}'
CANVAS_ASSIGNMENTS_URL: str = CANVAS_COURSE_URL + '/assignments'
CANVAS_ASSIGNMENT_URL: str = CANVAS_ASSIGNMENTS_URL + '/{aid}'
CANVAS_SUBMISSIONS_URL: str = CANVAS_ASSIGNMENT_URL + '/submissions'
CANVAS_SUBMISSION_URL: str = CANVAS_SUBMISSIONS_URL + '/{uid}'
CANVAS_UPDATE_GRADES_URL: str = CANVAS_SUBMISSIONS_URL + '/update_grades'

DEFAULT_SORT_BY: str = 'user_sortable_name'
DEFAULT_SELECT_COLUMNS: _tp.Tuple[str, ...] = (
    'user_id',
    'user_sortable_name',
    'grade',
    'score',
    'submitted_at',
    'preview_url',
    'attachments'
)

DEFAULT_OUT_FILEPATH: str = './assignment_data.xls'

DEFAULT_SUBMISSION_COMMENT_HEADER: str = 'submission_comment'
DEFAULT_GRADE_HEADER: str = 'grade'
DEFAULT_WORKBOOK_TAB: str = 'Sheet1'


DEFAULT_CHUNK_SIZE = 100

SLEEP_UPLOAD_APPRAISAL: float  = .02  # TODO Is this needed?
SLEEP_UPLOAD_APPRAISALS: float = 30.0  # TODO Seems high

del _tp
