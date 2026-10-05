"""Shared package cache for the launcher and the standalone updater."""
import hashlib
import os
from pathlib import Path
import re
import shutil
import tempfile
import zipfile


def _package_name(version):
    if not version or not re.fullmatch(r"\d+\.\d+\.\d+", version):
        raise ValueError("Invalid package version")
    return f"package_{version}.zip"


def _cache_paths(app_dir, version):
    name = _package_name(version)
    installation = os.path.normcase(os.path.realpath(app_dir))
    key = hashlib.sha256(installation.encode('utf-8')).hexdigest()[:16]
    local_data = Path(os.environ.get('LOCALAPPDATA') or Path.home() / 'AppData' / 'Local')
    return (
        local_data / 'Main_Program' / 'updates' / key / name,
        Path(app_dir) / '_internal' / 'updates' / name,
    )


def find_cached_package(app_dir, version):
    if not app_dir or not version:
        return None
    for path in _cache_paths(app_dir, version):
        try:
            if path.is_file() and zipfile.is_zipfile(path):
                return str(path)
        except OSError:
            continue
    return None


def store_cached_package(app_dir, version, source):
    """Publish a complete ZIP atomically; keep any existing cache on failure."""
    if not zipfile.is_zipfile(source):
        raise ValueError('Package is not a valid ZIP')
    failures = []
    for target in _cache_paths(app_dir, version):
        temporary = None
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile(dir=target.parent, suffix='.tmp', delete=False) as stream:
                temporary = Path(stream.name)
            shutil.copyfile(source, temporary)
            if not zipfile.is_zipfile(temporary):
                raise OSError('Copied package is not a valid ZIP')
            os.replace(temporary, target)
            return str(target)
        except OSError as error:
            failures.append(f'{target}: {error}')
        finally:
            if temporary is not None:
                try:
                    temporary.unlink(missing_ok=True)
                except OSError:
                    pass
    raise OSError('Cannot save update cache: ' + '; '.join(failures))
