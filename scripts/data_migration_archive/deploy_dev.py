#!/usr/bin/env python
"""
Deploy Offset ERP to the development server.

Usage (after you have backed up the database):
    python scripts/deploy_dev.py --confirm-backup

Windows shortcut:
    .\\deploy.ps1

Linux / macOS:
    ./deploy.sh

Options:
    --confirm-backup   Required. Confirms you already took a DB backup.
    --git-pull         Run `git pull` before deploying.
    --skip-pip         Skip pip install.
    --skip-preflight   Skip pre-migration data warnings.
    --run-tests        Run core test suites after migrate (slower).
    --dry-run          Print commands without running them.
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


class DeployError(Exception):
    pass


def _print_header(title: str) -> None:
    line = '=' * len(title)
    print(f'\n{title}\n{line}')


def run_command(cmd: list[str], *, dry_run: bool = False, check: bool = True) -> subprocess.CompletedProcess | None:
    display = ' '.join(cmd)
    print(f'>>> {display}')
    if dry_run:
        return None
    return subprocess.run(cmd, cwd=ROOT, check=check)


def ensure_project_root() -> None:
    if not (ROOT / 'manage.py').exists():
        raise DeployError(f'manage.py not found. Expected project root at {ROOT}')


def _model_has_field(model, field_name: str) -> bool:
    return field_name in {field.name for field in model._meta.get_fields()}


def run_preflight_checks(*, after_migrate: bool = False) -> None:
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
    import django

    django.setup()

    from django.db.models import Q

    from core.models import Dispatch
    from planning.models import PlanningJob

    title = 'Post-migration checks' if after_migrate else 'Pre-migration checks'
    _print_header(title)

    legacy_dispatch = Dispatch.objects.filter(Q(dc_no__isnull=True) | Q(dc_no='')).count()
    stage_reset = (
        PlanningJob.objects.filter(
            job_card__status__in=['released', 'in_production', 'completed', 'closed'],
        )
        .exclude(planning_stage__in=['', 'planning_done'])
        .count()
    )
    old_stage = PlanningJob.objects.filter(planning_stage='in_production').count()

    print(f'Dispatch rows that will get LEGACY-* DC numbers: {legacy_dispatch}')
    print(f'Released jobs whose planning sub-stage will move to planning_done: {stage_reset}')
    print(f'Jobs with old planning_stage=in_production (will rename): {old_stage}')

    qc_blocked = None
    if after_migrate and _model_has_field(PlanningJob, 'print_passes'):
        qc_blocked = PlanningJob.objects.filter(
            status__in=['pending_qc', 'draft'],
            print_passes__isnull=True,
        ).count()
        print(f'Jobs in draft/pending_qc missing print_passes: {qc_blocked}')

    warning_counts = [legacy_dispatch, stage_reset, old_stage]
    if qc_blocked is not None:
        warning_counts.append(qc_blocked)

    if any(warning_counts):
        print('\nReview these counts before using the updated app.')
    else:
        print('\nNo migration side-effects detected on current data.')


def show_pending_migrations(python_exe: str, *, dry_run: bool) -> None:
    _print_header('Migration status')
    run_command([python_exe, 'manage.py', 'showmigrations', '--plan'], dry_run=dry_run, check=False)


def maybe_collectstatic(python_exe: str, *, dry_run: bool) -> None:
    if dry_run:
        print('>>> collectstatic check skipped in dry-run mode')
        return

    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
    from django.conf import settings

    if not getattr(settings, 'STATIC_ROOT', None):
        print('>>> collectstatic skipped (STATIC_ROOT is not configured)')
        return

    _print_header('Static files')
    run_command([python_exe, 'manage.py', 'collectstatic', '--noinput'], dry_run=False)


def maybe_run_tests(python_exe: str, *, dry_run: bool) -> None:
    _print_header('Tests')
    run_command(
        [
            python_exe,
            'manage.py',
            'test',
            'planning.tests',
            'core.tests',
            'printing_plates.tests',
            'supply_chain.tests',
            '--verbosity=1',
        ],
        dry_run=dry_run,
    )


def repair_migrations(python_exe: str, *, dry_run: bool) -> None:
    _print_header('Migration repair')
    if dry_run:
        run_command([python_exe, 'scripts/data_migration_archive/repair_migrations.py'], dry_run=True, check=False)
        return
    run_command([python_exe, 'scripts/data_migration_archive/repair_migrations.py', '--apply'], dry_run=False, check=False)


def deploy(args: argparse.Namespace) -> None:
    ensure_project_root()
    python_exe = sys.executable

    if not args.confirm_backup:
        raise DeployError(
            'Database backup is required before deploy.\n'
            'Take your backup, then re-run with: --confirm-backup'
        )

    _print_header('Offset ERP - development deploy')
    print(f'Project root: {ROOT}')
    print(f'Python: {python_exe}')

    if args.git_pull:
        _print_header('Git pull')
        run_command(['git', 'pull'], dry_run=args.dry_run)

    if not args.skip_pip:
        _print_header('Python dependencies')
        run_command(
            [python_exe, '-m', 'pip', 'install', '-r', 'requirements.txt', '-q'],
            dry_run=args.dry_run,
        )

    _print_header('Django system check')
    run_command([python_exe, 'manage.py', 'check'], dry_run=args.dry_run)

    if not args.skip_preflight and not args.dry_run:
        run_preflight_checks(after_migrate=False)

    show_pending_migrations(python_exe, dry_run=args.dry_run)

    repair_migrations(python_exe, dry_run=args.dry_run)

    _print_header('Apply migrations')
    run_command([python_exe, 'manage.py', 'migrate', '--noinput'], dry_run=args.dry_run)

    if not args.skip_preflight and not args.dry_run:
        run_preflight_checks(after_migrate=True)

    maybe_collectstatic(python_exe, dry_run=args.dry_run)

    if args.run_tests:
        maybe_run_tests(python_exe, dry_run=args.dry_run)

    _print_header('Deploy complete')
    print('Next steps:')
    print('  1. Restart your app server / Windows service / gunicorn process.')
    print('  2. Open the site and smoke-test planning, production, and dispatch.')
    print('  3. Fill print_passes on any draft/pending_qc jobs if QC is blocked.')


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Deploy Offset ERP to the development server.')
    parser.add_argument(
        '--confirm-backup',
        action='store_true',
        help='Required. Confirm the database backup is already done.',
    )
    parser.add_argument('--git-pull', action='store_true', help='Run git pull before deploy.')
    parser.add_argument('--skip-pip', action='store_true', help='Skip pip install.')
    parser.add_argument('--skip-preflight', action='store_true', help='Skip pre-migration data warnings.')
    parser.add_argument('--run-tests', action='store_true', help='Run test suites after migrate.')
    parser.add_argument('--dry-run', action='store_true', help='Show commands without executing them.')
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    try:
        deploy(parse_args(argv))
    except DeployError as exc:
        print(f'\nDeploy aborted: {exc}', file=sys.stderr)
        return 1
    except subprocess.CalledProcessError as exc:
        print(f'\nDeploy failed: command exited with code {exc.returncode}', file=sys.stderr)
        return exc.returncode or 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
