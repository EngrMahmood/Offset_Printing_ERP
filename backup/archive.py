"""Backup archive helpers: 7-Zip (.7z, LZMA2) when available, plain .zip otherwise.

A .7z archive of the SQLite database is less than half the size of the same data in a
deflate .zip (about 8-9 MB against 19 MB for the production database). 7-Zip is a separate
program, so everything here degrades gracefully: if the binary is missing, or it fails,
the caller gets an ordinary .zip exactly as before.
"""
import logging
import os
import shutil
import subprocess
import tempfile
import zipfile

logger = logging.getLogger(__name__)

ARCHIVE_EXTENSIONS = ('.zip', '.7z')
_BINARIES = ('7zz', '7z', '7za', '7zr')
_WINDOWS_PATHS = (r'C:\Program Files\7-Zip\7z.exe', r'C:\Program Files (x86)\7-Zip\7z.exe')


class ArchiveError(Exception):
    pass


def find_7zip():
    """Path of a usable 7-Zip executable, or None."""
    for name in _BINARIES:
        found = shutil.which(name)
        if found:
            return found
    for path in _WINDOWS_PATHS:
        if os.path.isfile(path):
            return path
    return None


def is_7z(path):
    return str(path).lower().endswith('.7z')


def _run(command, cwd=None, timeout=3600):
    return subprocess.run(command, cwd=cwd, capture_output=True, text=True, timeout=timeout)


def create_7z(archive_path, db_path, db_name, level=5):
    """Create ``archive_path`` (.7z) holding ``db_path`` under the name ``db_name``.

    LZMA2 with a 32 MB dictionary (SQLite files repeat heavily, so a big dictionary pays off).
    The archive is verified with ``7z t`` before being trusted. Raises ArchiveError otherwise.
    """
    binary = find_7zip()
    if not binary:
        raise ArchiveError('7-Zip is not installed')
    if os.path.exists(archive_path):
        os.remove(archive_path)
    stage = tempfile.mkdtemp(prefix='backup_stage_', dir=os.path.dirname(archive_path) or None)
    try:
        staged = os.path.join(stage, db_name)
        try:
            os.link(db_path, staged)          # no 100 MB copy when the filesystem allows a hard link
        except OSError:
            shutil.copyfile(db_path, staged)
        result = _run(
            [binary, 'a', '-t7z', '-m0=lzma2', f'-mx={level}', '-md=32m', '-mmt=2', '-y', '-bso0', '-bsp0',
             os.path.abspath(archive_path), db_name],
            cwd=stage,
        )
        if result.returncode != 0:
            raise ArchiveError(f'7z failed ({result.returncode}): {(result.stderr or result.stdout)[-300:]}')
    finally:
        shutil.rmtree(stage, ignore_errors=True)
    verify = _run([binary, 't', '-y', '-bso0', '-bsp0', os.path.abspath(archive_path)])
    if verify.returncode != 0:
        os.remove(archive_path)
        raise ArchiveError(f'7z integrity test failed: {(verify.stderr or verify.stdout)[-300:]}')
    return archive_path


def list_entries(archive_path):
    """{name: uncompressed size} for a .zip or .7z archive."""
    if is_7z(archive_path):
        binary = find_7zip()
        if not binary:
            raise ArchiveError('This backup is a .7z archive but 7-Zip is not installed on this server.')
        result = _run([binary, 'l', '-slt', '-ba', os.path.abspath(archive_path)], timeout=600)
        if result.returncode != 0:
            raise ArchiveError(f'Cannot read the 7z archive: {(result.stderr or result.stdout)[-300:]}')
        entries, current = {}, {}
        for line in result.stdout.splitlines() + ['']:
            if not line.strip():
                if current.get('Path') and not current.get('Attributes', '').startswith('D'):
                    entries[current['Path']] = int(current.get('Size') or 0)
                current = {}
                continue
            key, _, value = line.partition(' = ')
            current[key.strip()] = value.strip()
        return entries
    with zipfile.ZipFile(archive_path, 'r') as archive:
        return {info.filename: info.file_size for info in archive.infolist() if not info.is_dir()}


def extract_all(archive_path, destination):
    os.makedirs(destination, exist_ok=True)
    if is_7z(archive_path):
        binary = find_7zip()
        if not binary:
            raise ArchiveError('This backup is a .7z archive but 7-Zip is not installed on this server.')
        result = _run([binary, 'x', '-y', f'-o{os.path.abspath(destination)}', os.path.abspath(archive_path)])
        if result.returncode != 0:
            raise ArchiveError(f'7z extraction failed: {(result.stderr or result.stdout)[-300:]}')
        return
    with zipfile.ZipFile(archive_path, 'r') as archive:
        archive.extractall(destination)
