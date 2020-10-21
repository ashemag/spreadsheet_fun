import gspread
from apiclient.discovery import build
from gspread.models import Spreadsheet
from gspread_dataframe import set_with_dataframe
from gspread_formatting import (
    set_column_width,
    set_row_height,
    set_frozen,
    cellFormat,
    format_cell_ranges,
    ConditionalFormatRule,
    GradientRule,
    CellFormat,
    get_conditional_format_rules,
    GridRange,
    textFormat,
    Color,
    GradientRule,
    InterpolationPoint,
)
import json
import os
import time
from oauth2client.service_account import ServiceAccountCredentials

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file',
]
WRITE_THRESHOLD = 1000

# colors
BLUE = {
    "red": 0.41,
    "green": 0.58,
    "blue": 0.93,
}
RED = {
    "red": 0.96,
    "green": 0.78,
    "blue": 0.77,
}
WHITE = {"green": 1.0, "red": 1.0, "blue": 1.0}
GREY = {"green": 0.41, "red": 0.41, "blue": 0.41}
GREEN = {"green": 0.73, "red": 0.34, "blue": 0.54}
COLOR_MAP = {'green': GREEN, "grey": GREY, "white": WHITE, "red": RED, "blue": BLUE}


def get_credentials(scopes=SCOPES, credential_password_env_var_name='GSHEET_PASSWORD'):
    """Read Google's JSON permission file.
    https://developers.google.com/api-client-library/python/auth/service-accounts#example
    :param scopes: List of scopes we need access to
    """
    cred_str = os.getenv(credential_password_env_var_name)
    cred_str = cred_str.replace("'", '"')
    cred_json = json.loads(cred_str, strict=False)

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(cred_json, SCOPES)
    gc = gspread.authorize(credentials)

    return credentials, gc


def grant_permissions(
    sheet_id,
    service_account_email,
    credential_password_env_var_name,
    share_domains=[],
    email_message="",
    send_notification=False,
    drive_api=None,
):
    if drive_api is None:
        credentials, _ = get_credentials(SCOPES, credential_password_env_var_name)
        drive_api = build('drive', 'v3', credentials=credentials)

    share_domains = list(set(share_domains + [service_account_email]))
    for domain in share_domains:
        domain_permission = {
            'type': 'user',
            'role': 'writer',
            'emailAddress': domain,
        }
        # exclude email message if we aren't sending notifications
        if not email_message or not send_notification:
            req = drive_api.permissions().create(
                fileId=sheet_id,
                body=domain_permission,
                fields="id",
                sendNotificationEmail=send_notification,
            )
        else:
            req = drive_api.permissions().create(
                fileId=sheet_id,
                body=domain_permission,
                emailMessage=email_message,
                fields="id",
                sendNotificationEmail=send_notification,
            )

        req.execute()


def create_google_spreadsheet(
    title,
    service_account_email,
    credential_password_env_var_name,
    email_message=None,
    send_notification=False,
    share_domains=None,
):
    """Create a new spreadsheet and open gspread object for it.
    .. note ::
        Created spreadsheet is not instantly visible in your Drive search and you need to access it by direct link.
    :param title: Spreadsheet title
    :param share_domains: List of Google Apps domain whose members get full access rights to the created sheet. Very handy, otherwise the file is visible only to the service worker itself. Example:: ``["redinnovation.com"]``.
    """

    credentials, gc = get_credentials(credential_password_env_var_name)

    drive_api = build('drive', 'v3', credentials=credentials)

    body = {
        'name': title,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
    }

    req = drive_api.files().create(body=body)
    new_sheet = req.execute()

    # Get id of fresh sheet
    sheet_id = new_sheet["id"]

    grant_permissions(
        sheet_id,
        service_account_email,
        credential_password_env_var_name,
        share_domains,
        email_message,
        send_notification=True,
        drive_api=drive_api,
    )

    link = f'\nhttps://docs.google.com/spreadsheets/d/{sheet_id}'
    return sheet_id, gc, link


def clear_sheet1(sh):
    # delete sheet1 if it exists
    try:
        wk = sh.worksheet('Sheet1')
        sh.del_worksheet(wk)
    except Exception as e:
        pass


def clear_worksheet(sh, title):
    try:
        worksheet = sh.add_worksheet(title=title, rows='500', cols='20')
    except:
        wk = sh.worksheet(title)
        wk.clear()
        sh.del_worksheet(wk)
        worksheet = sh.add_worksheet(title=title, rows='500', cols='20')


def conditional_formatter(wk, df, cols_to_colors, header_letters):
    rules = get_conditional_format_rules(wk)
    rules.clear()
    if cols_to_colors:
        for idx, col in enumerate(df.columns):
            if col in cols_to_colors:
                color = COLOR_MAP[cols_to_colors[col]]
            else:
                color = WHITE

            rule = ConditionalFormatRule(
                ranges=[
                    GridRange.from_a1_range(f'{header_letters[idx]}1:{header_letters[idx]}2000', wk)
                ],
                gradientRule=GradientRule(
                    minpoint=InterpolationPoint(color=WHITE, type="MIN"),
                    maxpoint=InterpolationPoint(color=color, type="MAX"),
                ),
            )
            rules.append(rule)

    rules.save()


def write_to_sheet(
    gc,
    sheet_id,
    worksheet_name,
    df,
    column_width_dict=None,
    header_row_height=None,
    no_borders=False,
    percent_cols=[],
    wrap_cols=[],
    cols_to_colors={},
):
    sh = gc.open_by_key(sheet_id)

    # may fail if sheet isn't created
    try:
        clear_worksheet(sh, worksheet_name)
    except Exception as _:
        pass

    wk = sh.worksheet(worksheet_name)
    header_letters = [
        'A',
        'B',
        'C',
        'D',
        'E',
        'F',
        'G',
        'H',
        'I',
        'J',
        'K',
        'L',
        'M',
        'N',
        'O',
        'P',
        'Q',
        'R',
        'S',
        'T',
        'U',
        'V',
        'W',
        'X',
        'Y',
        'Z',
    ]
    column_ids = header_letters.copy()

    # 26**2 columns max in spreadsheet
    for letter_A in header_letters:
        for letter_B in header_letters:
            column_ids.append(letter_A + letter_B)

    # writes df to sheet
    set_with_dataframe(wk, df)
    header_letter = column_ids[df.shape[1] - 1]

    # set all font to 11
    format_obj = {
        "textFormat": {"fontSize": 11},
    }

    # freeze the header layer
    set_frozen(wk, rows=1)

    # remove borders
    if no_borders:
        format_obj['borders'] = {
            "top": {"style": "NONE", "width": 0},
            "bottom": {"style": "NONE", "width": 0},
            "left": {"style": "NONE", "width": 0},
            "right": {"style": "NONE", "width": 0},
        }
        wk.format(f'A1:Z1000', format_obj)

    # set cols as percent cols
    for col in percent_cols:
        wk.format(
            f'{header_letters[col]}1:{header_letters[col]}2000',
            {'numberFormat': {'pattern': '0.0%', 'type': "PERCENT"}},
        )

    # writes a header
    wk.format(
        f'A1:{header_letter}1',
        {
            "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
            "horizontalAlignment": "CENTER",
            "textFormat": {"fontSize": 12, "bold": True},
            "wrapStrategy": "WRAP",
        },
    )

    # wrap columns
    for col in wrap_cols:
        wk.format(
            f'{header_letters[col]}1:{header_letters[col]}2000',
            {
                "wrapStrategy": "WRAP",
            },
        )

    # column widths
    for idx in range(df.shape[1]):
        if idx % WRITE_THRESHOLD == 0:
            time.sleep(2)
        if column_width_dict is not None and idx in column_width_dict:
            set_column_width(wk, column_ids[idx], column_width_dict[idx])
        else:
            set_column_width(wk, column_ids[idx], 150)

    time.sleep(2)

    # row heights
    if header_row_height:
        set_row_height(wk, '1', header_row_height)

    conditional_formatter(wk, df, cols_to_colors, header_letters)

    # clear sheet1 if it exists
    try:
        clear_sheet1(sh)
    except Exception as _:
        pass