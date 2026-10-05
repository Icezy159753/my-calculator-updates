"""Conservative component reuse and file-update release packaging for Windows CI."""
import argparse
import ast
import hashlib
import json
import os
import re
from pathlib import Path
import shutil
import subprocess
import sys
import urllib.error
import urllib.request
import zipfile

from release_files import digest, inventory, make_file_package, safe_name

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / 'build/release'
DIST = ROOT / 'dist'
TOOLCHAIN = 'windows-x64/python3.12/pyinstaller6.21.0/schema1'


def read_json(path, default=None):
    return json.loads(path.read_text(encoding='utf-8')) if path.exists() else default


def download(url, destination, authenticated=False):
    headers = {'User-Agent': 'Main-Program-Release'}
    if authenticated and os.environ.get('GITHUB_TOKEN'):
        headers['Authorization'] = 'Bearer ' + os.environ['GITHUB_TOKEN']
    request = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(request, timeout=90) as response, open(destination, 'wb') as output:
        shutil.copyfileobj(response, output)


def source_imports(path):
    tree = ast.parse(path.read_text(encoding='utf-8-sig'))
    return sorted({(node.module or '').split('.')[0] if isinstance(node, ast.ImportFrom)
                   else alias.name.split('.')[0]
                   for node in ast.walk(tree) if isinstance(node, (ast.Import, ast.ImportFrom))
                   for alias in node.names})


def fingerprints():
    def hash_inputs(paths, normalize_main=False):
        result = hashlib.sha256(TOOLCHAIN.encode())
        for path in sorted(paths):
            result.update(path.relative_to(ROOT).as_posix().encode())
            content = path.read_bytes()
            if normalize_main and path.name == 'Main_Program.py':
                tree = ast.parse(content.decode('utf-8-sig'))
                for node in tree.body:
                    if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'CURRENT_VERSION' for t in node.targets):
                        node.value = ast.Constant('VERSION_FROM_RELEASE_METADATA')
                content = ast.unparse(tree).encode('utf-8')
            result.update(content)
        return result.hexdigest()
    common = [ROOT / 'requirements_ttk.lock.txt', ROOT / '.github/workflows/windows-release.yml', ROOT / 'release_build.py']
    main_sources = [p for p in ROOT.glob('*.py') if not p.name.startswith(('test_', 'patch_'))
                    and p.name not in {'release_build.py', 'release_files.py', 'updater.py'}]
    program_imports = {p.relative_to(ROOT).as_posix(): source_imports(p)
                       for p in (ROOT / 'All_Programs').rglob('*.py') if p.name != '158_AutoLychee_OneFile.py'}
    runtime_files = [p for p in (ROOT / 'savReaderWriter').rglob('*') if p.is_file() and p.suffix.lower() in ('.py', '.dll', '.pyd')]
    icon = ROOT / 'Icon/I_Main.ico'
    main = hash_inputs([*common, ROOT / 'Main_Program.spec', ROOT / 'I_Main.ico', *main_sources,
                        *runtime_files, *([icon] if icon.exists() else [])], True)
    return {'main': hashlib.sha256((main + json.dumps(program_imports, sort_keys=True)).encode()).hexdigest(),
            'autolychee': hash_inputs([*common, ROOT / 'AutoLychee.spec', ROOT / 'All_Programs/158_AutoLychee_OneFile.py', ROOT / 'Icon/Autolychee.png']),
            'updater': hash_inputs([*common, ROOT / 'updater.py', ROOT / 'update_cache.py', ROOT / 'release_files.py', ROOT / 'setting.ico'])}


def extract_package(package, target):
    target.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(package) as archive:
        for entry in archive.infolist():
            name = entry.filename.rstrip('/')
            safe_name(name)
            path = target.joinpath(*name.split('/'))
            if entry.is_dir():
                path.mkdir(parents=True, exist_ok=True)
            else:
                path.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(entry) as source, open(path, 'wb') as output:
                    shutil.copyfileobj(source, output)
    if not (target / 'Main_Program.exe').is_file():
        raise RuntimeError('Previous full package has no launcher')


def prepare(repo, version):
    if not re.fullmatch(r'\d+\.\d+\.\d+', version or ''):
        raise ValueError('A numeric release version is required')
    BUILD.mkdir(parents=True, exist_ok=True)
    DIST.mkdir(exist_ok=True)
    inputs = fingerprints()
    plan = {'schema': 1, 'version': version, 'previous': None, 'inputs': inputs,
            'main': True, 'autolychee': True, 'updater': True}
    response_path = BUILD / 'previous-release.json'
    try:
        download(f'https://api.github.com/repos/{repo}/releases/latest', response_path, True)
        release = read_json(response_path)
        previous = release['tag_name'].lstrip('v')
        if tuple(map(int, previous.split('.'))) >= tuple(map(int, version.split('.'))):
            raise RuntimeError('Use a version newer than the published release')
        assets = {asset['name']: asset['browser_download_url'] for asset in release['assets']}
        full = assets.get(f'Main_Program_full_{previous}.zip') or assets.get('Main_Program.zip')
        if full:
            download(full, BUILD / 'previous.zip')
            extract_package(BUILD / 'previous.zip', BUILD / 'previous')
            plan['previous'] = previous
            old_manifest = None
            if 'release_manifest.json' in assets:
                download(assets['release_manifest.json'], BUILD / 'previous-manifest.json')
                old_manifest = read_json(BUILD / 'previous-manifest.json')
            old_files = inventory(BUILD / 'previous')
            (BUILD / 'previous-files.json').write_text(json.dumps(old_files), encoding='utf-8')
            # Reuse only an exact input match, with verified previous artifacts.
            if old_manifest and old_manifest.get('schema') == 1 and old_manifest.get('version') == previous:
                if old_files != old_manifest.get('files'):
                    raise RuntimeError('Previous package checksum manifest mismatch')
                for component in ('main', 'autolychee', 'updater'):
                    plan[component] = inputs[component] != old_manifest.get('inputs', {}).get(component)
            if not plan['updater']:
                download(assets['updater.exe'], DIST / 'updater.exe')
                if digest(DIST / 'updater.exe') != old_manifest.get('updater_sha256'):
                    raise RuntimeError('Previous updater checksum mismatch')
            if not plan['autolychee']:
                shutil.copytree(BUILD / 'previous/_internal/AutoLychee', DIST / 'AutoLychee', dirs_exist_ok=True)
            if not plan['main']:
                shutil.copytree(BUILD / 'previous', DIST / 'Main_Program', dirs_exist_ok=True)
    except urllib.error.HTTPError as error:
        if error.code != 404:
            raise
        print('No previous release; performing full build')
    (ROOT / 'release_version.json').write_text(json.dumps({'version': version}), encoding='utf-8')
    (BUILD / 'plan.json').write_text(json.dumps(plan), encoding='utf-8')
    if os.environ.get('GITHUB_OUTPUT'):
        with open(os.environ['GITHUB_OUTPUT'], 'a', encoding='utf-8') as output:
            for component in ('main', 'autolychee', 'updater'):
                output.write(f'{component}={str(plan[component]).lower()}\n')
    print(json.dumps({key: plan[key] for key in ('version', 'previous', 'main', 'autolychee', 'updater')}))


def assemble():
    plan = read_json(BUILD / 'plan.json')
    installed = DIST / 'Main_Program'
    if not plan['main']:
        # Replace bundled data as a unit; old runtime libraries remain intact.
        for folder in ('All_Programs', 'Icon'):
            target = installed / '_internal' / folder
            # These targets are fixed inside the build output, never user installations.
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(ROOT / folder, target, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
        if plan['autolychee']:
            target = installed / '_internal/AutoLychee'
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(DIST / 'AutoLychee', target)
        # Only explicitly declared package data is overlaid. Unknown runtime changes
        # are part of main source/spec fingerprints and require a full component build.
        for name in ('template.xlsx', 'Itemdef - Format.xlsx', 'Test3.json', 'openrouter.json'):
            if (ROOT / name).exists():
                shutil.copyfile(ROOT / name, installed / '_internal' / name)
    shutil.copyfile(ROOT / 'release_version.json', installed / '_internal/release_version.json')
    # Bytecode caches are generated locally, not release payloads.
    for path in installed.rglob('*.pyc'):
        if '__pycache__' in path.parts:
            path.unlink()


def package():
    plan = read_json(BUILD / 'plan.json')
    version = plan['version']
    installed = DIST / 'Main_Program'
    full = DIST / f'Main_Program_full_{version}.zip'
    with zipfile.ZipFile(full, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as archive:
        for path in sorted(installed.rglob('*')):
            if path.is_file():
                archive.write(path, path.relative_to(installed).as_posix())
    # Preserve the exact same ZIP bytes for legacy launchers/cache users.
    shutil.copyfile(full, DIST / 'Main_Program.zip')
    files = inventory(installed)
    if plan['previous']:
        delta = DIST / f"Main_Program_files_{plan['previous']}_to_{version}.zip"
        make_file_package(read_json(BUILD / 'previous-files.json'), installed, plan['previous'], version, delta)
        print(f'File update: {delta.stat().st_size} bytes; full: {full.stat().st_size} bytes')
    manifest = {'schema': 1, 'version': version, 'inputs': plan['inputs'], 'files': files,
                'updater_sha256': digest(DIST / 'updater.exe'), 'git_commit': subprocess.check_output(['git', 'rev-parse', 'HEAD'], text=True).strip()}
    (DIST / 'release_manifest.json').write_text(json.dumps(manifest, sort_keys=True), encoding='utf-8')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('command', choices=['prepare', 'assemble', 'package'])
    parser.add_argument('--repo')
    parser.add_argument('--version')
    args = parser.parse_args()
    if args.command == 'prepare':
        prepare(args.repo, args.version)
    elif args.command == 'assemble':
        assemble()
    else:
        package()
