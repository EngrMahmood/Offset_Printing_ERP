#!/usr/bin/env python
"""
Repair common migration history mismatches on existing databases.

Use when migrate fails with "table already exists" but Django still tries
to create it.

Usage:
    python scripts/repair_migrations.py
    python scripts/repair_migrations.py --apply
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


REPAIRS = (
    {
        'app': 'planning',
        'migration': '0030_jobcardchangerequest',
        'table': 'planning_jobcardchangerequest',
        'follow_up_columns': {},
    },
    {
        'app': 'planning',
        'migration': '0031_jobcardchangerequest_request_type_and_more',
        'table': 'planning_jobcardchangerequest',
        'follow_up_columns': {'request_type'},
    },
)


def django_setup() -> None:
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
    import django

    django.setup()


def migration_is_applied(app: str, migration_name: str) -> bool:
    from django.db import connection
    from django.db.migrations.recorder import MigrationRecorder

    recorder = MigrationRecorder(connection)
    return (app, migration_name) in recorder.applied_migrations()


def table_exists(table_name: str) -> bool:
    from django.db import connection

    return table_name in connection.introspection.table_names()


def table_columns(table_name: str) -> set[str]:
    from django.db import connection

    if not table_exists(table_name):
        return set()

    with connection.cursor() as cursor:
        description = connection.introspection.get_table_description(cursor, table_name)
    return {column.name for column in description}


def should_fake_repair(repair: dict) -> tuple[bool, str]:
    app = repair['app']
    migration = repair['migration']
    table = repair['table']
    required_columns = repair.get('follow_up_columns') or set()

    if migration_is_applied(app, migration):
        return False, f'{app}.{migration} already applied'

    if not table_exists(table):
        return False, f'{table} does not exist yet'

    if required_columns:
        existing_columns = table_columns(table)
        if not required_columns.issubset(existing_columns):
            missing = ', '.join(sorted(required_columns - existing_columns))
            return False, f'{table} is missing columns for {app}.{migration}: {missing}'

    return True, f'{table} exists but {app}.{migration} is not recorded'


def find_repairs() -> list[tuple[dict, str]]:
    matches = []
    for repair in REPAIRS:
        should_fake, reason = should_fake_repair(repair)
        if should_fake:
            matches.append((repair, reason))
    return matches


def fake_migration(python_exe: str, app: str, migration: str) -> None:
    cmd = [python_exe, 'manage.py', 'migrate', app, migration, '--fake']
    print(f'>>> {" ".join(cmd)}')
    subprocess.run(cmd, cwd=ROOT, check=True)


def repair(*, apply: bool) -> int:
    django_setup()
    matches = find_repairs()

    if not matches:
        print('No migration repairs needed.')
        return 0

    print('Migration repairs detected:')
    for repair, reason in matches:
        print(f'  - {repair["app"]}.{repair["migration"]}: {reason}')

    if not apply:
        print('\nDry run only. Re-run with --apply to mark these migrations as applied.')
        return 0

    python_exe = sys.executable
    for repair, _reason in matches:
        fake_migration(python_exe, repair['app'], repair['migration'])

    print('\nRepair complete. Now run: python manage.py migrate')
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Repair migration history mismatches.')
    parser.add_argument(
        '--apply',
        action='store_true',
        help='Mark detected migrations as applied with --fake.',
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    try:
        return repair(apply=parse_args(argv).apply)
    except subprocess.CalledProcessError as exc:
        print(f'Repair failed with exit code {exc.returncode}', file=sys.stderr)
        return exc.returncode or 1


if __name__ == '__main__':
    raise SystemExit(main())
