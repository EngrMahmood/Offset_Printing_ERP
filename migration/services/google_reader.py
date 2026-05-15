import json
import logging
import os
import csv
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from urllib.request import Request as UrlRequest, urlopen
import io

from django.conf import settings
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow

logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
OAUTH_CLIENT_FILE_CANDIDATES = [
    'google_oauth_client_secret.json',
    'google-oauth-client-secret.json',
    'client_secret.json',
    'credentials.json',
]


def _get_gspread_config_dir():
    appdata = os.getenv('APPDATA')
    if appdata:
        return Path(appdata) / 'gspread'
    return Path.home() / '.config' / 'gspread'


DEFAULT_GSPREAD_CREDENTIALS_FILE = _get_gspread_config_dir() / 'credentials.json'
DEFAULT_GSPREAD_AUTHORIZED_USER_FILE = _get_gspread_config_dir() / 'authorized_user.json'


def _get_config_value(name):
    value = os.getenv(name)
    if value is not None:
        return value
    return getattr(settings, name, None)


def _load_service_account_info():
    json_text = _get_config_value('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not json_text:
        return None

    try:
        return json.loads(json_text)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            'Google service account JSON is configured but contains invalid JSON.'
        ) from exc


def _load_json_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as handle:
            return json.load(handle)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f'Google OAuth client config file contains invalid JSON: {file_path}') from exc


def _get_oauth_client_file_candidates():
    configured_path = _get_config_value('GOOGLE_OAUTH_CLIENT_SECRETS_FILE')
    if configured_path:
        return [configured_path]

    base_dir = Path(getattr(settings, 'BASE_DIR', Path.cwd()))
    candidates = [str(base_dir / filename) for filename in OAUTH_CLIENT_FILE_CANDIDATES]
    candidates.append(str(DEFAULT_GSPREAD_CREDENTIALS_FILE))

    for pattern in ('client_secret*.json', '*oauth*client*.json', 'credentials*.json'):
        for matched_path in base_dir.glob(pattern):
            candidates.append(str(matched_path))

    return list(dict.fromkeys(candidates))


def _load_oauth_client_config():
    inline_json = _get_config_value('GOOGLE_OAUTH_CLIENT_CONFIG_JSON')
    if inline_json:
        try:
            return json.loads(inline_json), 'GOOGLE_OAUTH_CLIENT_CONFIG_JSON'
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                'Google OAuth client config JSON is configured but contains invalid JSON.'
            ) from exc

    client_id = _get_config_value('GOOGLE_OAUTH_CLIENT_ID')
    client_secret = _get_config_value('GOOGLE_OAUTH_CLIENT_SECRET')
    if client_id and client_secret:
        return {
            'web': {
                'client_id': client_id,
                'client_secret': client_secret,
                'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
                'token_uri': 'https://oauth2.googleapis.com/token',
            }
        }, 'GOOGLE_OAUTH_CLIENT_ID/SECRET'

    for candidate in _get_oauth_client_file_candidates():
        if os.path.isfile(candidate):
            return _load_json_file(candidate), candidate

    return None, None


def is_google_oauth_configured():
    client_config, _ = _load_oauth_client_config()
    if not client_config:
        return False

    return bool(client_config.get('web') or client_config.get('installed'))


def _load_default_oauth_credentials():
    if DEFAULT_GSPREAD_AUTHORIZED_USER_FILE.is_file():
        return Credentials.from_authorized_user_file(str(DEFAULT_GSPREAD_AUTHORIZED_USER_FILE), scopes=SCOPES)
    return None


def _extract_sheet_identifiers(sheet_url):
    parsed = urlparse(sheet_url)
    path_parts = [part for part in parsed.path.split('/') if part]

    spreadsheet_id = ''
    gid = ''
    if 'd' in path_parts:
        index = path_parts.index('d')
        if index + 1 < len(path_parts):
            spreadsheet_id = path_parts[index + 1]

    query = parse_qs(parsed.query)
    if query.get('gid'):
        gid = query['gid'][0]

    if not gid and parsed.fragment:
        fragment_query = parse_qs(parsed.fragment)
        if fragment_query.get('gid'):
            gid = fragment_query['gid'][0]
        elif 'gid=' in parsed.fragment:
            gid = parsed.fragment.split('gid=', 1)[1].split('&', 1)[0]

    return spreadsheet_id, gid or None


import re


def _normalize_header_row(rows):
    if not rows:
        return rows

    first_row = rows[0]
    numeric_keys = all(
        (key == '' or re.match(r'^[0-9]+(\.[0-9]+)?$', key))
        for key in first_row.keys()
    )
    if not numeric_keys:
        return rows

    if len(rows) < 2:
        return rows

    header_values = list(first_row.values())
    if not header_values or any(val is None or str(val).strip() == '' for val in header_values):
        return rows

    header_names = [str(val).strip().lower() for val in header_values]
    known_headers = {
        'po', 'sku', 'customer', 'quantity', 'order_qty', 'destination',
        'jc', 'job', 'job name', 'job_name', 'job card', 'job_card',
        'job no', 'jobnumber', 'job_number', 'jc_number', 'job_card_no',
    }
    if not any(any(known in name for known in known_headers) for name in header_names):
        return rows

    normalized_rows = []
    for row in rows[1:]:
        normalized = {}
        for key, header in zip(first_row.keys(), header_names):
            normalized[header] = row.get(key)
        normalized_rows.append(normalized)
    return normalized_rows


def _clean_field_name(key):
    cleaned = re.sub(r'[^0-9a-z]+', '_', str(key).strip().lower())
    return cleaned.strip('_')


def _looks_like_numeric_header_line(line):
    try:
        row = next(csv.reader([line]))
    except Exception:
        return False
    if not row:
        return False
    numeric_values = 0
    non_empty_values = 0
    for token in row:
        token = token.strip()
        if token == '':
            continue
        non_empty_values += 1
        if re.match(r'^[0-9]+(\.[0-9]+)?$', token):
            numeric_values += 1
    if non_empty_values == 0:
        return False
    return numeric_values >= max(1, non_empty_values - 2)


def _read_csv_rows_from_url(url):
    request = UrlRequest(url, headers={'User-Agent': 'Mozilla/5.0'})

    try:
        with urlopen(request, timeout=20) as response:
            content_type = response.headers.get('Content-Type', '')
            payload = response.read().decode('utf-8-sig')
    except Exception:
        return None

    if 'text/html' in content_type.lower() or payload.lstrip().startswith('<!DOCTYPE html') or payload.lstrip().startswith('<html'):
        return None

    lines = payload.splitlines()
    if lines and _looks_like_numeric_header_line(lines[0]):
        payload = '\n'.join(lines[1:])

    rows = list(csv.DictReader(io.StringIO(payload)))
    if not rows:
        return []

    rows = _normalize_header_row(rows)

    normalized_rows = []
    for row in rows:
        normalized = {}
        for key, value in row.items():
            normalized[_clean_field_name(key)] = value
        normalized_rows.append(normalized)
    return normalized_rows


def _read_public_google_sheet(sheet_url):
    spreadsheet_id, gid = _extract_sheet_identifiers(sheet_url)
    if not spreadsheet_id:
        return None

    candidate_urls = []
    if gid:
        candidate_urls.append(f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}')
    candidate_urls.append(f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv')
    candidate_urls.append(f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:csv')

    for candidate_url in candidate_urls:
        rows = _read_csv_rows_from_url(candidate_url)
        if rows is not None:
            return rows

    return None


def read_google_sheet_metadata(sheet_url, oauth_token=None, sample_size=5):
    rows = _read_public_google_sheet(sheet_url)
    if rows is not None:
        columns = list(rows[0].keys()) if rows else []
        sample_rows = rows[:sample_size]
        return {
            'sheet_url': sheet_url,
            'columns': columns,
            'sample_rows': sample_rows,
        }

    try:
        import gspread

        if oauth_token:
            credentials = Credentials(**oauth_token)
        else:
            credentials = _load_default_oauth_credentials()

        if not credentials:
            raise RuntimeError('No Google credentials were found for private sheet access.')

        client = gspread.authorize(credentials)
        spreadsheet = client.open_by_url(sheet_url)
        worksheet = spreadsheet.sheet1
        values = worksheet.get_all_values()
        if not values:
            return {'sheet_url': sheet_url, 'columns': [], 'sample_rows': []}

        headers = [str(value).strip() for value in values[0]]
        data_rows = []
        for row in values[1:1 + sample_size]:
            normalized_row = {headers[idx]: row[idx] if idx < len(row) else '' for idx in range(len(headers))}
            data_rows.append(normalized_row)

        return {
            'sheet_url': sheet_url,
            'columns': headers,
            'sample_rows': data_rows,
        }
    except Exception as exc:
        raise RuntimeError(f'Could not read Google Sheet metadata: {exc}') from exc


def get_google_credential_status():
    service_account_json = _get_config_value('GOOGLE_SERVICE_ACCOUNT_JSON')
    if service_account_json:
        try:
            json.loads(service_account_json)
            return {
                'available': True,
                'method': 'GOOGLE_SERVICE_ACCOUNT_JSON',
                'message': 'Google service account credentials are configured via environment variable.',
            }
        except json.JSONDecodeError:
            return {
                'available': False,
                'method': 'GOOGLE_SERVICE_ACCOUNT_JSON',
                'message': 'GOOGLE_SERVICE_ACCOUNT_JSON is set but contains invalid JSON.',
            }

    credentials_path = _get_config_value('GOOGLE_APPLICATION_CREDENTIALS')
    if credentials_path:
        if os.path.isfile(credentials_path):
            return {
                'available': True,
                'method': 'GOOGLE_APPLICATION_CREDENTIALS',
                'message': 'Google service account credentials file is available.',
            }
        return {
            'available': False,
            'method': 'GOOGLE_APPLICATION_CREDENTIALS',
            'message': f'Google credentials file not found at {credentials_path}.',
        }

    if DEFAULT_GSPREAD_AUTHORIZED_USER_FILE.is_file():
        return {
            'available': True,
            'method': 'gspread_authorized_user',
            'message': f'Google OAuth user credentials are available at {DEFAULT_GSPREAD_AUTHORIZED_USER_FILE}.',
        }

    default_paths = [
        os.path.expanduser('~/.config/gspread/service_account.json'),
        os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'gspread', 'service_account.json'),
    ]
    for default_path in default_paths:
        if default_path and os.path.isfile(default_path):
            return {
                'available': True,
                'method': 'gspread_default_path',
                'message': f'Google credentials found at default path: {default_path}',
            }

    if is_google_oauth_configured():
        _, source = _load_oauth_client_config()
        return {
            'available': False,
            'method': 'google_oauth',
            'message': (
                'No service account credentials were found, but Google OAuth is configured. '
                f'Google OAuth client configuration is available from {source}. '
                'The app will redirect you to Google to grant access.'
            ),
        }

    return {
        'available': False,
        'method': None,
        'message': (
            'No Google private-sheet credentials are configured. '
            'Simplest setup: make the Google Sheet viewable by anyone with the link and paste the URL here. '
            'If the sheet must stay private, configure Google OAuth or a service-account credential file.'
        ),
    }


def has_google_sheet_access(oauth_token=None):
    if oauth_token:
        return True

    credential_status = get_google_credential_status()
    return credential_status['available']


def has_configured_google_sheet_auth(oauth_token=None):
    if oauth_token:
        return True

    if _load_service_account_info():
        return True

    credentials_path = _get_config_value('GOOGLE_APPLICATION_CREDENTIALS')
    if credentials_path and os.path.isfile(credentials_path):
        return True

    if _load_default_oauth_credentials():
        return True

    default_paths = [
        os.path.expanduser('~/.config/gspread/service_account.json'),
        os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'gspread', 'service_account.json'),
    ]
    return any(default_path and os.path.isfile(default_path) for default_path in default_paths)


def _credentials_from_token(token):
    return Credentials(
        token=token.get('token'),
        refresh_token=token.get('refresh_token'),
        token_uri=token.get('token_uri'),
        client_id=token.get('client_id'),
        client_secret=token.get('client_secret'),
        scopes=token.get('scopes'),
    )


def _credentials_to_dict(credentials):
    return {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes,
    }


def credentials_to_dict(credentials):
    return _credentials_to_dict(credentials)


def get_google_oauth_flow(request, redirect_uri):
    client_config, _ = _load_oauth_client_config()
    if not client_config:
        raise RuntimeError(
            'Google OAuth client is not configured. Set GOOGLE_OAUTH_CLIENT_ID and GOOGLE_OAUTH_CLIENT_SECRET, '
            'set GOOGLE_OAUTH_CLIENT_CONFIG_JSON, or add a Google OAuth client JSON file such as client_secret.json in the project root.'
        )

    config_key = 'web' if client_config.get('web') else 'installed'
    if config_key == 'installed':
        client_config = {'web': dict(client_config['installed'])}
    else:
        client_config = {'web': dict(client_config['web'])}

    client_config['web']['redirect_uris'] = [redirect_uri]
    client_config['web'].setdefault('auth_uri', 'https://accounts.google.com/o/oauth2/auth')
    client_config['web'].setdefault('token_uri', 'https://oauth2.googleapis.com/token')
    flow = Flow.from_client_config(client_config, scopes=SCOPES)
    flow.redirect_uri = redirect_uri
    return flow


def read_google_sheet(sheet_url, oauth_token=None):
    """Read a Google Sheet URL and return row dicts.

    Supports credentials via:
    - GOOGLE_APPLICATION_CREDENTIALS file path
    - GOOGLE_SERVICE_ACCOUNT_JSON JSON payload
    - Google OAuth credentials stored in session
    - gspread default auth behavior
    """
    try:
        import gspread
    except ImportError as exc:
        raise RuntimeError('gspread is not installed. Add gspread to environment first.') from exc

    public_rows = _read_public_google_sheet(sheet_url)
    if public_rows is not None:
        return public_rows

    if not has_configured_google_sheet_auth(oauth_token=oauth_token):
        raise RuntimeError(
            'This Google Sheet is not publicly readable yet. Make it viewable by anyone with the link, '
            'or configure Google OAuth/service-account credentials for private sheet access.'
        )

    service_account_info = _load_service_account_info()
    credentials_path = _get_config_value('GOOGLE_APPLICATION_CREDENTIALS')
    default_oauth_credentials = _load_default_oauth_credentials()
    try:
        if oauth_token:
            credentials = _credentials_from_token(oauth_token)
            if credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
                oauth_token.update(_credentials_to_dict(credentials))
            client = gspread.authorize(credentials)
        elif default_oauth_credentials:
            if default_oauth_credentials.expired and default_oauth_credentials.refresh_token:
                default_oauth_credentials.refresh(Request())
            client = gspread.authorize(default_oauth_credentials)
        elif service_account_info:
            client = gspread.service_account_from_dict(service_account_info)
        elif credentials_path:
            client = gspread.service_account(filename=credentials_path)
        else:
            client = gspread.service_account()

        spreadsheet = client.open_by_url(sheet_url)
        worksheet = spreadsheet.sheet1
        rows = worksheet.get_all_records()
    except FileNotFoundError as exc:
        raise RuntimeError(
            'Google credentials file was not found. Set GOOGLE_APPLICATION_CREDENTIALS to a valid service account JSON file path, '
            'or set GOOGLE_SERVICE_ACCOUNT_JSON with the JSON contents, or make the sheet viewable by anyone with the link.'
        ) from exc
    except Exception as exc:
        logger.exception('Failed to read Google Sheet: %s', sheet_url)
        raise RuntimeError(
            'Unable to read Google Sheet. Verify the sheet URL is correct, the sheet is viewable by anyone with the link or shared with the authorized Google account, and the credentials are configured.'
        ) from exc

    normalized_rows = []
    for row in rows:
        normalized = {}
        for key, value in row.items():
            normalized[_clean_field_name(key)] = value
        normalized_rows.append(normalized)
    return normalized_rows
