"""Verified file updates with a durable rollback journal (no GUI/network dependencies)."""
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import tempfile
import zipfile

PROTECTED = {'Test3.json', 'openrouter.json', 'Itemdef - Format.xlsx',
             '_internal/Test3.json', '_internal/openrouter.json', '_internal/Itemdef - Format.xlsx'}


class UpdateRejected(ValueError):
    pass


class RollbackFailed(RuntimeError):
    pass


def digest(path):
    value = hashlib.sha256()
    with open(path, 'rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            value.update(chunk)
    return value.hexdigest()


def safe_name(name):
    if not isinstance(name, str) or not name or '\\' in name or ':' in name:
        raise UpdateRejected('Invalid file path')
    parts = name.split('/')
    reserved = {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(1, 10)), *(f'LPT{i}' for i in range(1, 10))}
    if PurePosixPath(name).is_absolute() or any(
        p in ('', '.', '..') or p.rstrip(' .') != p or p.split('.')[0].upper() in reserved
        or any(ord(c) < 32 or c in '<>"|?*' for c in p) for p in parts
    ):
        raise UpdateRejected('Unsafe file path')
    return name


def protected(name):
    lower = name.casefold()
    return (lower in {p.casefold() for p in PROTECTED}
            or lower.startswith(('savreaderwriter/', '_internal/updates/'))
            or lower.startswith('_internal/update-transactions/')
            or lower in {'updater.exe', 'updater.lock', 'update_debug.log', 'updater_debug.log'})


def destination(root, name):
    safe_name(name)
    root = Path(root).resolve()
    path = root.joinpath(*name.split('/'))
    if not path.resolve().is_relative_to(root):
        raise UpdateRejected('Path leaves installation')
    for candidate in (path, *path.parents):
        if candidate == root:
            break
        if candidate.is_symlink() or (hasattr(candidate, 'is_junction') and candidate.is_junction()):
            raise UpdateRejected('Linked installation path')
    return path


def _write_json(path, data):
    temporary = path.with_suffix('.tmp')
    with open(temporary, 'w', encoding='utf-8') as stream:
        json.dump(data, stream, ensure_ascii=False)
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, path)


def _rollback(root, transaction, journal):
    try:
        for entry in reversed(journal['entries']):
            if protected(entry['path']):
                raise UpdateRejected('Protected rollback path')
            target = destination(root, entry['path'])
            if entry['before'] is None:
                if target.exists():
                    target.unlink()
            else:
                backup = destination(transaction / 'backup', entry['path'])
                if digest(backup) != entry['before']:
                    raise UpdateRejected('Rollback backup checksum mismatch')
                target.parent.mkdir(parents=True, exist_ok=True)
                temporary = target.with_name(target.name + '.rollback-tmp')
                shutil.copyfile(backup, temporary)
                os.replace(temporary, target)
        _write_json(transaction / 'journal.json', {**journal, 'state': 'rolled-back'})
    except Exception as error:
        raise RollbackFailed(f'Rollback incomplete; backup retained at {transaction}') from error


def recover_pending(root):
    transactions = destination(root, '_internal/update-transactions')
    if not transactions.exists():
        return
    for transaction in transactions.iterdir():
        if transaction.is_symlink() or (hasattr(transaction, 'is_junction') and transaction.is_junction()):
            raise UpdateRejected('Linked transaction directory')
        journal_path = transaction / 'journal.json'
        if not transaction.is_dir() or not journal_path.exists():
            continue
        journal = json.loads(journal_path.read_text(encoding='utf-8'))
        if journal.get('state') not in ('prepared', 'committed', 'rolled-back', 'applying'):
            raise UpdateRejected('Invalid recovery journal')
        if journal.get('state') == 'applying':
            _rollback(root, transaction, journal)
        if journal.get('state') in ('prepared', 'committed', 'rolled-back', 'applying'):
            shutil.rmtree(transaction)


def apply_file_update(root, package, current_version, target_version):
    """Validate before changing files. Restore originals on failure or next startup."""
    recover_pending(root)
    transactions = destination(root, '_internal/update-transactions')
    transactions.mkdir(parents=True, exist_ok=True)
    transaction = Path(tempfile.mkdtemp(prefix='update-', dir=transactions))
    applying = False
    journal = None
    try:
        with zipfile.ZipFile(package) as archive:
            if archive.getinfo('manifest.json').file_size > 8 * 1024 * 1024:
                raise UpdateRejected('Manifest too large')
            manifest = json.loads(archive.read('manifest.json'))
            if manifest.get('schema') != 1 or manifest.get('from') != current_version or manifest.get('to') != target_version:
                raise UpdateRejected('Update version mismatch')
            entries = manifest['files']
            if not entries or len(entries) > 50000:
                raise UpdateRejected('Invalid file list')
            names = set()
            expected_members = {'manifest.json'}
            for entry in entries:
                name = safe_name(entry['path'])
                if name.casefold() in names or protected(name) or name.casefold().startswith('_internal/update-transactions/'):
                    raise UpdateRejected('Duplicate or protected path')
                names.add(name.casefold())
                target = destination(root, name)
                before, after = entry.get('before'), entry.get('after')
                if before is None and after is None:
                    raise UpdateRejected('Empty change')
                for checksum in (before, after):
                    if checksum is not None and not re.fullmatch('[0-9a-f]{64}', checksum):
                        raise UpdateRejected('Invalid checksum')
                if before is None:
                    if target.exists():
                        raise UpdateRejected(f'Unexpected existing file: {name}')
                elif not target.is_file() or digest(target) != before:
                    raise UpdateRejected(f'Installed file differs: {name}')
                if after is not None:
                    member = 'payload/' + name
                    expected_members.add(member)
                    size = entry['size']
                    if not isinstance(size, int) or size < 0 or archive.getinfo(member).file_size != size:
                        raise UpdateRejected('Payload size mismatch')
                    stage = destination(transaction / 'stage', name)
                    stage.parent.mkdir(parents=True, exist_ok=True)
                    with archive.open(member) as source, open(stage, 'wb') as output:
                        shutil.copyfileobj(source, output)
                    if digest(stage) != after:
                        raise UpdateRejected(f'Payload checksum mismatch: {name}')
            members = archive.namelist()
            if len(members) != len(set(members)) or set(members) != expected_members:
                raise UpdateRejected('Unexpected ZIP members')
        # Backup everything before any installed file is changed.
        for entry in entries:
            if entry.get('before') is not None:
                backup = destination(transaction / 'backup', entry['path'])
                backup.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(destination(root, entry['path']), backup)
                if digest(backup) != entry['before']:
                    raise UpdateRejected('Installed file changed during preparation')
        entries = sorted(entries, key=lambda e: e['path'] == '_internal/release_version.json')
        journal = {'state': 'applying', 'entries': entries}
        _write_json(transaction / 'journal.json', journal)
        applying = True
        for entry in entries:
            target = destination(root, entry['path'])
            if entry.get('after') is None:
                target.unlink()
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                os.replace(destination(transaction / 'stage', entry['path']), target)
        _write_json(transaction / 'journal.json', {**journal, 'state': 'committed'})
        applying = False
    except Exception:
        if applying:
            _rollback(root, transaction, journal)
            applying = False
        raise
    finally:
        if not applying:
            shutil.rmtree(transaction, ignore_errors=True)


def inventory(root):
    return {p.relative_to(root).as_posix(): {'sha256': digest(p), 'size': p.stat().st_size}
            for p in sorted(Path(root).rglob('*')) if p.is_file() and not protected(p.relative_to(root).as_posix())}


def make_file_package(old_files, new_root, from_version, to_version, output):
    new_files = inventory(new_root)
    changes = []
    for name in sorted(set(old_files) | set(new_files)):
        if protected(name):
            continue
        before = old_files.get(name, {}).get('sha256')
        after = new_files.get(name, {}).get('sha256')
        if before != after:
            changes.append({'path': name, 'before': before, 'after': after,
                            'size': new_files.get(name, {}).get('size', 0)})
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as archive:
        archive.writestr('manifest.json', json.dumps({'schema': 1, 'from': from_version, 'to': to_version, 'files': changes}))
        for entry in changes:
            if entry['after'] is not None:
                archive.write(Path(new_root) / entry['path'], 'payload/' + entry['path'])
    return new_files


def extract_full_package(package, target):
    """Reject dangerous/ambiguous paths before extracting a full release."""
    with zipfile.ZipFile(package) as archive:
        names = set()
        for info in archive.infolist():
            name = safe_name(info.filename.rstrip('/'))
            if name.casefold() in names or (info.external_attr >> 16) & 0o170000 == 0o120000:
                raise UpdateRejected('Duplicate or linked ZIP member')
            names.add(name.casefold())
            destination(target, name)
        for info in archive.infolist():
            path = destination(target, info.filename.rstrip('/'))
            if info.is_dir():
                path.mkdir(parents=True, exist_ok=True)
            else:
                path.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(info) as source, open(path, 'wb') as output:
                    shutil.copyfileobj(source, output)


def install_full_update(root, source, current_version, target_version, work_dir):
    """Repair/upgrade managed release files transactionally; keep all user extras."""
    recover_pending(root)
    new_files = inventory(source)
    old_files = {}
    for name in new_files:
        path = destination(root, name)
        if path.is_file():
            old_files[name] = {'sha256': digest(path), 'size': path.stat().st_size}
    if all(old_files.get(name) == info for name, info in new_files.items()):
        return
    package = Path(work_dir) / 'full-install-transaction.zip'
    make_file_package(old_files, source, current_version, target_version, package)
    try:
        apply_file_update(root, package, current_version, target_version)
    finally:
        package.unlink(missing_ok=True)
