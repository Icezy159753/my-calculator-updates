"""Double-click deployment: version, release gates, review, atomic push, CI watch."""
import argparse
import hashlib
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
import json
if __package__:
    from .build_preflight import run_preflight
else:
    from build_preflight import run_preflight

ROOT = Path(__file__).resolve().parents[1]
REPO = 'Icezy159753/my-calculator-updates'
VERSION = re.compile(r'(?m)^(CURRENT_VERSION\s*=\s*")(\d+\.\d+\.\d+)(")')
TESTS = ['test_autolychee_packaging', 'test_update_cache', 'test_release_files', 'test_release_build',
         'tools.test_deploy_release', 'tools.test_build_preflight']


def run(args, capture=False, timeout=None, env=None, cwd=None):
    process = subprocess.Popen(args, cwd=cwd or ROOT, env=env, stdin=subprocess.PIPE,
                               stdout=subprocess.PIPE if capture else None,
                               stderr=subprocess.PIPE if capture else None)
    try:
        output, error = process.communicate(b'', timeout=timeout)
    except subprocess.TimeoutExpired:
        if os.name == 'nt':
            subprocess.run(['taskkill', '/F', '/T', '/PID', str(process.pid)], capture_output=True)
        else:
            process.kill()
        process.communicate()
        raise RuntimeError('คำสั่งทดสอบเกินเวลา จึงหยุดการเผยแพร่') from None
    if process.returncode:
        if capture and error:
            print(error.decode('utf-8', errors='replace'))
        raise RuntimeError(f'คำสั่งไม่ผ่าน (exit {process.returncode}): {args[0]} {args[1]}')
    if not capture:
        return ''
    decoded = output.decode('utf-8', errors='replace')
    return decoded if '\0' in decoded else decoded.strip()


def git(*args):
    return run(['git', *args], capture=True)


def version_tuple(value):
    if not re.fullmatch(r'\d+\.\d+\.\d+', value):
        raise ValueError('เลขเวอร์ชันต้องเป็น x.y.z')
    return tuple(map(int, value.split('.')))


def next_version(current, tags):
    published = [version_tuple(tag[1:]) for tag in tags if re.fullmatch(r'v\d+\.\d+\.\d+', tag)]
    highest = max(published, default=(0, 0, 0))
    value = version_tuple(current)
    if value <= highest:
        value = (*highest[:2], highest[2] + 1)
    return '.'.join(map(str, value))


def release_input(name):
    path = Path(name)
    if path.name.casefold() in {'test3.json', 'openrouter.json', '.env'}:
        return False
    if name.startswith(('All_Programs/', 'Icon/', 'hooks/', '.github/workflows/')):
        return '__pycache__' not in path.parts and path.suffix.lower() in {
            '.py', '.png', '.jpg', '.svg', '.ico', '.ttf', '.wav', '.yml', '.yaml'}
    if name in {'tools/deploy_release.py', 'tools/test_deploy_release.py', 'tools/build_preflight.py',
                'tools/test_build_preflight.py', 'Deploy_GitHub.bat',
                'template.xlsx', 'Itemdef - Format.xlsx'}:
        return True
    return len(path.parts) == 1 and (path.suffix.lower() in {'.py', '.spec', '.ico', '.md'}
                                    or path.name.startswith('requirements')
                                    or ('Build exe Update' in name and path.suffix == '.txt'))


def changes():
    tracked = set(filter(None, git('diff', '--name-only', '-z', 'HEAD').split('\0')))
    untracked = set(filter(None, git('ls-files', '--others', '--exclude-standard', '-z').split('\0')))
    all_changes = tracked | untracked
    selected = sorted(name for name in all_changes if release_input(name))
    omitted = sorted(all_changes - set(selected))
    staged = set(filter(None, git('diff', '--cached', '--name-only', '-z').split('\0')))
    if staged - set(selected):
        raise RuntimeError('มีไฟล์นอกชุด release ถูก stage อยู่ กรุณา unstage ไฟล์เหล่านั้นก่อน')
    return selected, omitted


def snapshot(names):
    return {name: hashlib.sha256((ROOT / name).read_bytes()).hexdigest()
            if (ROOT / name).is_file() else None for name in names}


def release_gates():
    print('\nกำลังทดสอบ release (หากไม่ผ่านจะไม่ commit/push)...', flush=True)
    run_preflight(ROOT)
    env = os.environ.copy()
    env['QT_QPA_PLATFORM'] = 'offscreen'
    (ROOT / 'build').mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='deploy-smoke-', dir=ROOT / 'build') as data:
        env['AUTOLYCHEE_DATA'] = data
        run([sys.executable, '-X', 'utf8', '-m', 'unittest', *TESTS, '-q'], timeout=180, env=env)
        for args, marker, timeout in [
            (['All_Programs/158_AutoLychee_OneFile.py', '--smoke-test'], 'Auto Lychee GUI smoke test OK', 120),
            (['All_Programs/158_AutoLychee_OneFile.py', '--post'], None, 60),
            (['Main_Program.py', '--run-module', '158_AutoLychee_OneFile', '--entry-point', 'run_this_app', '--check'], '12 sections OK', 120),
        ]:
            output = run([sys.executable, '-X', 'utf8', *args], capture=True, timeout=timeout, env=env)
            if marker and marker not in output:
                raise RuntimeError(f'ไม่พบผลตรวจ {marker}; ดู AUTOLYCHEE_RELEASE.md')
            print(f'  {args[-1]} ผ่าน', flush=True)


def watch_release(tag, commit):
    print('\nกำลังรอ GitHub Actions เริ่ม...', flush=True)
    for _ in range(18):
        runs = json.loads(run(['gh', 'run', 'list', '--repo', REPO, '--workflow', 'windows-release.yml',
                              '--branch', tag, '--commit', commit, '-L', '1', '--json', 'databaseId,url'], capture=True))
        if runs:
            print(runs[0]['url'], flush=True)
            run(['gh', 'run', 'watch', str(runs[0]['databaseId']), '--repo', REPO, '--exit-status'])
            return
        time.sleep(5)
    raise RuntimeError('push สำเร็จแล้ว แต่ยังไม่พบงาน build กรุณาตรวจ GitHub Actions; อย่าลบหรือย้าย tag')


def deploy(options):
    written = {}
    staging_started = False
    try:
        if getattr(options, 'check_only', False):
            release_gates()
            print('\nตรวจผ่านแล้ว ไม่มีการเพิ่มเวอร์ชัน/commit/push')
            return
        if git('branch', '--show-current') != 'main':
            raise RuntimeError('กรุณาใช้ branch main ก่อน deploy')
        remote = git('remote', 'get-url', 'origin').removesuffix('.git')
        if remote not in {f'https://github.com/{REPO}', f'git@github.com:{REPO}', f'ssh://git@github.com/{REPO}'}:
            raise RuntimeError('origin ไม่ใช่ repository ของ Main_Program')
        if not options.dry_run:
            if not shutil.which('gh'):
                raise RuntimeError('ต้องติดตั้ง GitHub CLI (gh) และใช้ gh auth login ก่อน')
            run(['gh', 'auth', 'status'], capture=True)
            git('fetch', 'origin', 'main', '--tags')
        ahead, behind = map(int, git('rev-list', '--left-right', '--count', 'HEAD...origin/main').split())
        if behind:
            raise RuntimeError('เครื่องนี้ตามหลัง origin/main กรุณารวมโค้ดล่าสุดก่อน deploy (สคริปต์ไม่แก้ไฟล์งานให้เอง)')
        selected, omitted = changes()
        source_path = ROOT / 'Main_Program.py'
        source = source_path.read_text(encoding='utf-8-sig')
        match = VERSION.search(source)
        if not match:
            raise RuntimeError('ไม่พบ CURRENT_VERSION ใน Main_Program.py')
        current = match[2]
        tags = git('tag', '-l', 'v*').splitlines()
        retry = False
        if not options.dry_run and not selected and ahead and 'v' + current in tags:
            retry = (git('rev-parse', 'v' + current) == git('rev-parse', 'HEAD')
                     and not git('ls-remote', '--tags', 'origin', 'refs/tags/v' + current))
        if not selected and not ahead and not options.version_only:
            print('ยังไม่มีไฟล์เปลี่ยน จึงไม่สร้าง release ใหม่\nถ้าต้องการทดสอบเลขเวอร์ชัน: Deploy_GitHub.bat --version-only')
            return
        version = current if retry else next_version(current, tags)
        tag = 'v' + version
        print(f'\n{"ส่งซ้ำ release ที่ push ไม่สำเร็จ" if retry else "เตรียม release"}: {tag}')
        print(f'  commit ที่ยังไม่ขึ้น origin/main: {ahead}')
        for name in selected:
            print('  ส่ง: ' + name)
        for name in omitted:
            print('  ไม่เพิ่มเข้า release: ' + name)
        if options.dry_run:
            print('\nDry run: ใช้ tags ในเครื่องเท่านั้น ไม่มีการแก้ไฟล์ ทดสอบ commit หรือ push')
            return
        if not retry:
            instructions = list(ROOT.glob('*Build exe Update*Github.txt'))
            if len(instructions) != 1:
                raise RuntimeError('ต้องมีไฟล์คำสั่ง Build exe Update ขึ้น Github.txt หนึ่งไฟล์')
            for path, text in [(source_path, VERSION.sub(lambda m: m[1] + version + m[3], source, count=1)),
                               (instructions[0], re.sub(r'\d+\.\d+\.\d+', version, instructions[0].read_text(encoding='utf-8-sig')))]:
                original = path.read_bytes()
                path.write_text(text, encoding='utf-8')
                written[path] = (original, path.read_bytes())
        selected, _ = changes()
        tested = snapshot(selected)
        release_gates()
        print(f'\nผ่านการทดสอบแล้ว จะเผยแพร่ {tag} พร้อมไฟล์ข้างต้น')
        if not options.yes:
            answer = input('กด Enter เพื่อเผยแพร่ หรือพิมพ์ N เพื่อยกเลิก: ').strip().lower()
            if answer not in ('', 'y', 'yes'):
                print('ยกเลิกแล้ว')
                return
        latest, _ = changes()
        if latest != selected or snapshot(latest) != tested:
            raise RuntimeError('ไฟล์มีการแก้ไขระหว่างทดสอบ กรุณารันใหม่เพื่อทดสอบชุดล่าสุด')
        if not retry:
            staging_started = True
            if selected:
                run(['git', 'add', '-A', '--', *selected])
                run(['git', 'commit', '-m', f'Release {tag}'])
            git('tag', tag)
        # Both refs are sent together; auto-tag sees the already existing tag.
        print(f'\nส่ง main และ {tag} ขึ้น GitHub...', flush=True)
        try:
            run(['git', 'push', '--atomic', 'origin', 'main', tag])
        except Exception:
            print(f'commit/tag ยังอยู่ในเครื่อง รัน BAT อีกครั้งเพื่อส่งซ้ำ หรือใช้: git push --atomic origin main {tag}')
            raise
        staging_started = True
        if not options.no_watch:
            watch_release(tag, git('rev-parse', 'HEAD'))
        print(f'\nส่งขึ้น GitHub แล้ว: https://github.com/{REPO}/releases/tag/{tag}')
    finally:
        # Restore only our own version edits before staging, preserving IDE edits.
        if not staging_started:
            for path, (original, ours) in written.items():
                if path.read_bytes() == ours:
                    path.write_bytes(original)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='เผยแพร่ Main_Program ขึ้น GitHub อัตโนมัติ')
    parser.add_argument('--dry-run', action='store_true', help='แสดงแผนจาก tags ในเครื่อง โดยไม่เปลี่ยนหรือส่งไฟล์')
    parser.add_argument('--check-only', action='store_true', help='ตรวจ preflight/tests/smoke โดยไม่เพิ่มเวอร์ชันหรือส่ง GitHub')
    parser.add_argument('--version-only', action='store_true', help='สร้างรุ่นทดสอบได้แม้ไม่มีโปรแกรมเปลี่ยน')
    parser.add_argument('--yes', action='store_true', help='ไม่ถามยืนยันหลังทดสอบ')
    parser.add_argument('--no-watch', action='store_true', help='ส่ง tag แล้วไม่รอผล build')
    try:
        deploy(parser.parse_args())
    except (Exception, KeyboardInterrupt) as error:
        print('\nหยุด deploy: ' + str(error))
        sys.exit(1)
