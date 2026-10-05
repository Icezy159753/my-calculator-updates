"""CI gate: install the delta over the actual previous release, then smoke test."""
import json
import os
from pathlib import Path
import subprocess
from release_build import BUILD, DIST, ROOT
from release_files import apply_file_update, inventory, digest


def verify():
    process = subprocess.Popen([str(DIST / 'updater.exe'), '--self-test'], stdin=subprocess.PIPE,
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    try:
        stdout, stderr = process.communicate(b'', timeout=120)
    except subprocess.TimeoutExpired:
        subprocess.run(['taskkill', '/F', '/T', '/PID', str(process.pid)])
        raise
    if process.returncode or b'Updater transaction self-test OK' not in stdout:
        raise RuntimeError('Frozen updater self-test failed: ' + stderr.decode('utf-8', errors='replace'))
    programs = ROOT / 'All_Programs'
    for source in programs.glob('*.py'):
        bundled = DIST / 'Main_Program/_internal/All_Programs' / source.name
        if not bundled.is_file() or digest(bundled) != digest(source):
            raise RuntimeError(f'Program missing or stale in package: {source.name}')
        compile(source.read_text(encoding='utf-8-sig'), str(source), 'exec')
    plan = json.loads((BUILD / 'plan.json').read_text(encoding='utf-8'))
    if not plan['previous']:
        print('Initial release: no previous installation to upgrade')
        return
    package = DIST / f"Main_Program_files_{plan['previous']}_to_{plan['version']}.zip"
    previous = BUILD / 'previous'
    apply_file_update(previous, package, plan['previous'], plan['version'])
    manifest = json.loads((DIST / 'release_manifest.json').read_text(encoding='utf-8'))
    if inventory(previous) != manifest['files']:
        raise RuntimeError('Upgraded installation differs from the new full release')
    env = os.environ.copy()
    env['QT_QPA_PLATFORM'] = 'offscreen'
    env['AUTOLYCHEE_DATA'] = str(BUILD / 'transition-smoke')
    env['PYINSTALLER_RESET_ENVIRONMENT'] = '1'
    for executable, args in [
        (previous / '_internal/AutoLychee/AutoLychee.exe', ['--smoke-test']),
        (previous / '_internal/AutoLychee/AutoLychee.exe', ['--post']),
        (previous / 'Main_Program.exe', ['--run-module', '158_AutoLychee_OneFile', '--entry-point', 'run_this_app', '--check']),
    ]:
        process = subprocess.Popen([str(executable), *args], env=env, stdin=subprocess.PIPE,
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        try:
            stdout, stderr = process.communicate(b'', timeout=120)
        except subprocess.TimeoutExpired:
            subprocess.run(['taskkill', '/F', '/T', '/PID', str(process.pid)])
            raise
        if process.returncode or ('--smoke-test' in args and b'Auto Lychee GUI smoke test OK' not in stdout):
            raise RuntimeError(f'Transition smoke failed: {executable.name}; {stderr.decode("utf-8", errors="replace")}')
    print('Previous installed version upgraded successfully; all managed files and smoke tests verified')


if __name__ == '__main__':
    verify()
