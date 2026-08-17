import json
import logging
import os

from django.conf import settings

logger = logging.getLogger(__name__)

WRITE_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


class SheetsAuthError(Exception):
    """Credentials are missing, invalid, or not shared with the spreadsheet."""


def _get_config_value(name):
    value = os.getenv(name)
    if value is not None:
        return value
    return getattr(settings, name, None)


def _default_credentials_path():
    appdata = os.getenv('APPDATA')
    base = appdata if appdata else os.path.join(os.path.expanduser('~'), '.config')
    return os.path.join(base, 'gspread', 'service_account.json')


def get_gspread_client():
    """Build an authenticated gspread client for writing, using (in order):

    1. GOOGLE_SHEETS_SERVICE_ACCOUNT_JSON env var (inline JSON payload)
    2. GOOGLE_SERVICE_ACCOUNT_JSON env var (shared with the read-only migration importer)
    3. SheetsSyncSetting.service_account_json_path
    4. GOOGLE_APPLICATION_CREDENTIALS env var
    5. the default gspread credentials path (%APPDATA%/gspread/service_account.json)
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise SheetsAuthError('gspread is not installed. Add it to requirements.txt.') from exc

    from sheets_sync.models import SheetsSyncSetting

    inline_json = _get_config_value('GOOGLE_SHEETS_SERVICE_ACCOUNT_JSON') or _get_config_value('GOOGLE_SERVICE_ACCOUNT_JSON')
    if inline_json:
        try:
            info = json.loads(inline_json)
        except json.JSONDecodeError as exc:
            raise SheetsAuthError('Google service account JSON env var contains invalid JSON.') from exc
        credentials = Credentials.from_service_account_info(info, scopes=WRITE_SCOPES)
        return gspread.authorize(credentials)

    configured_path = SheetsSyncSetting.get_settings().service_account_json_path
    candidate_path = configured_path or _get_config_value('GOOGLE_APPLICATION_CREDENTIALS') or _default_credentials_path()

    if not os.path.isfile(candidate_path):
        raise SheetsAuthError(
            f'Google service-account credentials file not found at "{candidate_path}". '
            'Configure SheetsSyncSetting.service_account_json_path, or set '
            'GOOGLE_SHEETS_SERVICE_ACCOUNT_JSON / GOOGLE_APPLICATION_CREDENTIALS.'
        )

    credentials = Credentials.from_service_account_file(candidate_path, scopes=WRITE_SCOPES)
    return gspread.authorize(credentials)


def open_spreadsheet():
    from sheets_sync.models import SheetsSyncSetting

    spreadsheet_id = SheetsSyncSetting.get_settings().spreadsheet_id
    if not spreadsheet_id:
        raise SheetsAuthError('SheetsSyncSetting.spreadsheet_id is not configured.')

    client = get_gspread_client()
    try:
        return client.open_by_key(spreadsheet_id)
    except Exception as exc:
        raise SheetsAuthError(
            f'Could not open spreadsheet "{spreadsheet_id}". Verify the ID is correct and the '
            'spreadsheet is shared (Editor access) with the service account email.'
        ) from exc


def get_or_create_worksheet(spreadsheet, tab_name, header_row):
    try:
        worksheet = spreadsheet.worksheet(tab_name)
    except Exception:
        worksheet = spreadsheet.add_worksheet(title=tab_name, rows=1000, cols=max(len(header_row), 10))
        worksheet.update('A1', [header_row])
        return worksheet

    existing_header = worksheet.row_values(1)
    if existing_header != header_row:
        worksheet.update('A1', [header_row])
    return worksheet
