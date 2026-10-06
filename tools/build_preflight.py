"""Read-only build checks; never import/execute application programs or spec files."""
import ast
from dataclasses import dataclass
import importlib.metadata
import io
from pathlib import Path
import re
import subprocess
import sys
import tarfile


@dataclass(frozen=True)
class Import:
    module: str
    line: int
    optional: bool = False
    relative: int = 0
    dynamic: bool = False


def parse(source, filename):
    # Same grammar target as windows-release.yml; this is not a full 3.12 execution.
    tree = ast.parse(source, filename=filename, feature_version=(3, 12))
    compile(tree, filename, 'exec')  # Also catch return/await/break outside their scopes.
    return tree


def imports(tree):
    result = []
    class Visitor(ast.NodeVisitor):
        optional = False
        def visit_Try(self, node):
            previous = self.optional
            guarded = any(handler.type is None or any(
                isinstance(n, ast.Name) and n.id in {'ImportError', 'ModuleNotFoundError', 'Exception'}
                for n in ast.walk(handler.type)) for handler in node.handlers)
            self.optional = previous or guarded
            for item in node.body: self.visit(item)
            self.optional = previous
            for item in [*node.handlers, *node.orelse, *node.finalbody]: self.visit(item)
        def visit_If(self, node):
            if isinstance(node.test, ast.Name) and node.test.id == 'TYPE_CHECKING':
                for item in node.orelse: self.visit(item)
                return
            if isinstance(node.test, ast.Attribute) and node.test.attr == 'TYPE_CHECKING':
                for item in node.orelse: self.visit(item)
                return
            self.generic_visit(node)
        def visit_Import(self, node):
            result.extend(Import(alias.name, node.lineno, self.optional) for alias in node.names)
        def visit_ImportFrom(self, node):
            if node.module:
                result.append(Import(node.module, node.lineno, self.optional, node.level))
            elif node.level:
                result.extend(Import(alias.name, node.lineno, self.optional, node.level) for alias in node.names)
        def visit_Call(self, node):
            function = node.func
            dynamic = ((isinstance(function, ast.Name) and function.id == '__import__')
                       or (isinstance(function, ast.Attribute) and function.attr == 'import_module'))
            if dynamic and node.args and isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str):
                value = node.args[0].value
                if not value.startswith('.'):
                    result.append(Import(value, node.lineno, self.optional, dynamic=True))
            elif dynamic:
                result.append(Import('<dynamic expression>', node.lineno, self.optional, dynamic=True))
            self.generic_visit(node)
    Visitor().visit(tree)
    return result


def normalize(value):
    return re.sub(r'[-_.]+', '-', value).lower()


def ci_packages(root):
    declared = {name: set() for name in ('main', 'autolychee', 'updater')}
    text = (root / '.github/workflows/windows-release.yml').read_text(encoding='utf-8-sig')
    scope = set(declared)
    for line in text.splitlines():
        components = re.findall(r'steps\.plan\.outputs\.(main|autolychee|updater)', line)
        if components: scope = set(components)
        match = re.search(r'python -m pip install\s+(.+)', line)
        if not match: continue
        packages = []
        args = match[1]
        for filename in re.findall(r'-r\s+["\']?([\w./-]+)', args):
            path = root / filename
            if not path.is_file(): raise ValueError(f'ไม่พบ requirements ที่ workflow ใช้: {filename}')
            packages += path.read_text(encoding='utf-8-sig').splitlines()
        args = re.sub(r'-r\s+["\']?[\w./-]+["\']?', '', args)
        packages += args.split()
        for value in packages:
            match = re.match(r'^([A-Za-z0-9][A-Za-z0-9_.-]*)(?:\[.*?\])?(?:[=<>!~;\s]|$)', value.strip())
            if match:
                for component in scope: declared[component].add(normalize(match[1]))
    return declared


ALIASES = {'PIL': 'Pillow', 'cv2': 'opencv-python', 'sklearn': 'scikit-learn', 'docx': 'python-docx',
           'yaml': 'PyYAML', 'bs4': 'beautifulsoup4', 'pythoncom': 'pywin32', 'pywintypes': 'pywin32',
           'win32com': 'pywin32', 'win32api': 'pywin32', 'win32gui': 'pywin32', 'win32con': 'pywin32',
           'win32job': 'pywin32', 'win32process': 'pywin32', 'win32clipboard': 'pywin32',
           'googleapiclient': 'google-api-python-client', 'google_auth_oauthlib': 'google-auth-oauthlib',
           'google.auth': 'google-auth', 'google.oauth2': 'google-auth', 'google.generativeai': 'google-generativeai',
           'google.genai': 'google-genai', 'distutils': 'setuptools'}


def distributions(module, mapping):
    for alias in sorted(ALIASES, key=len, reverse=True):
        if module == alias or module.startswith(alias + '.'):
            return {normalize(ALIASES[alias])}
    root = module.split('.')[0]
    return {normalize(value) for value in mapping.get(root, [])} or {normalize(root)}


def spec_info(tree):
    hidden, collected, excluded = set(), set(), set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Assign, ast.AugAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            if any(isinstance(target, ast.Name) and target.id == 'hiddenimports' for target in targets):
                if isinstance(node.value, (ast.List, ast.Tuple)):
                    hidden.update(n.value for n in node.value.elts if isinstance(n, ast.Constant) and isinstance(n.value, str))
        if isinstance(node, ast.Call):
            if isinstance(node.func, ast.Name) and node.func.id in {'collect_all', 'collect_submodules'}:
                if node.args and isinstance(node.args[0], ast.Constant): collected.add(node.args[0].value)
            for keyword in node.keywords:
                if keyword.arg == 'excludes' and isinstance(keyword.value, (ast.List, ast.Tuple)):
                    excluded.update(n.value for n in keyword.value.elts if isinstance(n, ast.Constant))
    return hidden, collected, excluded


def local_file(root, filename, entry):
    bits = entry.module.split('.')
    starts = [root, root / 'All_Programs']
    if entry.relative:
        start = root / filename
        start = start.parent
        for _ in range(entry.relative - 1): start = start.parent
        starts = [start]
    for start in starts:
        path = start.joinpath(*bits)
        if path.with_suffix('.py').is_file() or (path / '__init__.py').is_file(): return True
        # Namespace packages bundled as data are valid too.
        if path.is_dir() and any(path.glob('*.py')): return True
    return False


def baseline_local(filename, entry, baseline):
    starts = [Path('.'), Path('All_Programs')]
    if entry.relative:
        start = Path(filename).parent
        for _ in range(entry.relative - 1): start = start.parent
        starts = [start]
    for start in starts:
        path = start.joinpath(*entry.module.split('.'))
        if path.with_suffix('.py').as_posix() in baseline or (path / '__init__.py').as_posix() in baseline:
            return True
    return False


def baseline_sources(root):
    listing = subprocess.run(['git', 'ls-tree', '--name-only', '-z', 'origin/main'], cwd=root, capture_output=True, timeout=30)
    if listing.returncode: raise RuntimeError('อ่าน origin/main ไม่ได้ กรุณา git fetch origin main ก่อน')
    root_sources = [name for name in listing.stdout.decode('utf-8').split('\0') if name.endswith('.py')
                    and not name.startswith(('test_', 'patch_'))]
    args = ['git', 'archive', 'origin/main', '--', 'All_Programs', 'Main_Program.spec', 'AutoLychee.spec', *root_sources]
    result = subprocess.run(args, cwd=root, capture_output=True, timeout=30)
    if result.returncode: raise RuntimeError('อ่าน source ของ origin/main ไม่ได้ กรุณา git fetch origin main ก่อน')
    if len(result.stdout) > 64 * 1024 * 1024: raise RuntimeError('baseline source ใหญ่เกินขอบเขตตรวจ')
    with tarfile.open(fileobj=io.BytesIO(result.stdout)) as archive:
        return {member.name: archive.extractfile(member).read().decode('utf-8-sig')
                for member in archive.getmembers() if member.isfile() and member.name.endswith(('.py', '.spec'))}


def check(root, baseline, mapping=None):
    root = Path(root)
    errors, warnings, trees = [], [], {}
    paths = set((root / 'All_Programs').rglob('*.py')) | set((root / 'hooks').rglob('*.py'))
    paths |= {path for path in root.glob('*.py') if not path.name.startswith(('test_', 'patch_'))}
    paths |= {root / name for name in ('Main_Program.py', 'updater.py', 'release_build.py', 'release_files.py',
                                      'update_cache.py', 'verify_file_transition.py', 'Main_Program.spec', 'AutoLychee.spec')}
    for path in sorted(paths):
        name = path.relative_to(root).as_posix()
        if not path.is_file():
            errors.append(f'{name}: ไม่พบไฟล์ที่ release ต้องใช้')
            continue
        try: trees[name] = parse(path.read_text(encoding='utf-8-sig'), name)
        except (SyntaxError, UnicodeError, ValueError) as error:
            errors.append(f'{name}:{getattr(error, "lineno", 1)}: syntax/encoding ใช้กับ CI Python 3.12 ไม่ได้ ({getattr(error, "msg", type(error).__name__)})')
    try: declared = ci_packages(root)
    except (OSError, ValueError) as error:
        errors.append(str(error)); declared = {component: set() for component in ('main', 'autolychee', 'updater')}
    mapping = importlib.metadata.packages_distributions() if mapping is None else mapping
    specs = {component: spec_info(trees.get(filename, ast.Module(body=[], type_ignores=[])))
             for component, filename in [('main', 'Main_Program.spec'), ('autolychee', 'AutoLychee.spec')]}
    specs['updater'] = (set(), set(), set())
    for component, filename in [('main', 'Main_Program.spec'), ('autolychee', 'AutoLychee.spec')]:
        try:
            old_hidden, old_collected, _ = spec_info(parse(baseline.get(filename, ''), filename))
        except (SyntaxError, ValueError): old_hidden, old_collected = set(), set()
        hidden, collected, _ = specs[component]
        for module in (hidden - old_hidden) | (collected - old_collected):
            if not isinstance(module, str) or module.split('.')[0] in sys.stdlib_module_names: continue
            if local_file(root, filename, Import(module, 1)): continue
            if not distributions(module, mapping) & declared[component]:
                warnings.append(f'{filename}: {module} — hiddenimports/collect ใหม่ แต่ CI ยังไม่ได้ติดตั้ง package ที่รองรับ')
    old = {}
    for name, source in baseline.items():
        if name.endswith(('.py', '.spec')):
            try: old[name] = imports(parse(source, name))
            except (SyntaxError, ValueError): pass
    direct = {component: set() for component in specs}
    for name, tree in trees.items():
        if name == 'Main_Program.py': direct['main'].update(i.module for i in imports(tree) if not i.relative)
        if name == 'updater.py': direct['updater'].update(i.module for i in imports(tree) if not i.relative)
    for name, tree in trees.items():
        build_only = name.endswith('.spec') or name.startswith('hooks/')
        component = 'autolychee' if name.endswith('158_AutoLychee_OneFile.py') or name == 'AutoLychee.spec' else ('updater' if name == 'updater.py' else 'main')
        previous = {(i.module, i.relative, i.optional, i.dynamic) for i in old.get(name, [])}
        hidden, collected, excluded = specs[component]
        # Already used by programs in the published component, rather than
        # blindly assuming every optional dependency is part of the runtime.
        known = {i.module for old_name, entries in old.items() for i in entries if not i.optional
                 and (('autolychee' if old_name.endswith('158_AutoLychee_OneFile.py') else
                       'updater' if old_name == 'updater.py' else 'main') == component)}
        sections = set(re.findall(r'^# ====== MODULE:\s*(\w+)', (root / name).read_text(encoding='utf-8-sig'), re.M)) if component == 'autolychee' else set()
        for entry in imports(tree):
            root_module = entry.module.split('.')[0]
            fresh = (entry.module, entry.relative, entry.optional, entry.dynamic) not in previous
            if baseline_local(name, entry, baseline) and not local_file(root, name, entry):
                (warnings if entry.optional else errors).append(f'{name}:{entry.line}: {entry.module} — โมดูล local ที่เคยมีถูกลบ/เปลี่ยนชื่อ แต่ยังมีการ import')
            if not fresh: continue
            location = f'{name}:{entry.line}: {entry.module}'
            severity = warnings if entry.optional else errors
            if entry.module == '<dynamic expression>':
                warnings.append(location + ' — ชื่อโมดูลคำนวณตอนรัน ตรวจ hiddenimports/hook เพิ่มเติม'); continue
            if root_module in sections: continue
            if local_file(root, name, entry):
                if (not entry.relative and (root / (root_module + '.py')).is_file()
                        and entry.module not in hidden and entry.module not in direct[component]
                        and name.startswith('All_Programs/')):
                    warnings.append(location + ' — โมดูล local นอก All_Programs; ตรวจ hiddenimports/การ bundle source ให้ครบ')
                continue
            if entry.relative:
                severity.append(location + ' — ไม่พบโมดูล relative ที่อ้างถึง'); continue
            if root_module in excluded and not build_only:
                severity.append(location + f' — {component} spec excludes โมดูลนี้ ต้องแยก EXE/แก้เส้นทางเรียก'); continue
            if root_module in sys.stdlib_module_names: continue
            candidates = distributions(entry.module, mapping)
            if not candidates & declared[component]:
                severity.append(location + f' — CI ส่วน {component} ไม่ได้ติดตั้ง package ที่รองรับ ({", ".join(sorted(candidates))}); เพิ่ม requirements/workflow ก่อน')
                continue
            covered = (entry.module in hidden or entry.module in known or entry.module in direct[component]
                       or any(entry.module == prefix or entry.module.startswith(prefix + '.') for prefix in collected))
            if not covered and component != 'updater' and not build_only:
                warnings.append(location + ' — เป็น import ใหม่ ยังไม่ยืนยัน hiddenimports/hook ของ spec; ตรวจว่า EXE รวมโมดูลนี้ด้วย')
            if entry.dynamic and not covered:
                warnings.append(location + ' — dynamic import อาจต้องเพิ่ม hiddenimports')
    # Literal data and icons must exist in the repository; CI-generated files are exceptions.
    generated = {'Test3.json', 'openrouter.json', 'release_version.json'}
    for name in ('Main_Program.spec', 'AutoLychee.spec'):
        if name not in trees: continue
        bindings = {target.id: node.value.value for node in trees[name].body if isinstance(node, ast.Assign)
                    and isinstance(node.value, ast.Constant) for target in node.targets if isinstance(target, ast.Name)}
        for node in ast.walk(trees[name]):
            items = []
            if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'datas' for t in node.targets):
                if isinstance(node.value, (ast.List, ast.Tuple)): items = list(node.value.elts)
            if isinstance(node, ast.Call):
                if isinstance(node.func, ast.Name) and node.func.id == 'Analysis' and node.args and isinstance(node.args[0], (ast.List, ast.Tuple)):
                    items += [ast.copy_location(ast.Tuple(elts=[value], ctx=ast.Load()), value) for value in node.args[0].elts]
                for keyword in node.keywords:
                    if keyword.arg == 'icon': items += [ast.copy_location(ast.Tuple(elts=[keyword.value], ctx=ast.Load()), keyword.value)]
                    if keyword.arg == 'datas' and isinstance(keyword.value, (ast.List, ast.Tuple)): items += list(keyword.value.elts)
            for item in items:
                if isinstance(item, ast.Tuple) and item.elts and isinstance(item.elts[0], (ast.Constant, ast.Name)):
                    value = item.elts[0].value if isinstance(item.elts[0], ast.Constant) else bindings.get(item.elts[0].id)
                    if not isinstance(value, str) or value in generated: continue
                    if not (root / value).exists(): errors.append(f'{name}:{item.lineno}: ไม่พบไฟล์/โฟลเดอร์สำหรับ bundle: {value}')
    return sorted(set(errors)), sorted(set(warnings)), len(trees)


def run_preflight(root):
    print('\nตรวจความพร้อม build EXE ก่อนเผยแพร่...', flush=True)
    errors, warnings, count = check(root, baseline_sources(root))
    for message in warnings: print('  WARNING: ' + message)
    for message in errors: print('  ERROR: ' + message)
    if errors: raise RuntimeError(f'พบ {len(errors)} ปัญหาที่ต้องแก้ก่อนเผยแพร่ — ยังไม่ commit/push')
    print(f'  ผ่านการตรวจ {count} ไฟล์; warnings={len(warnings)}', flush=True)
    return warnings
