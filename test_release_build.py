import ast
import os
from pathlib import Path
import shutil
import json
from types import SimpleNamespace
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from unittest.mock import Mock
import sys
import hashlib
import io
import zipfile
from release_files import inventory, make_file_package, apply_file_update

import release_build as build


class ReleaseBuildTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        paths = ['requirements_ttk.lock.txt', '.github/workflows/windows-release.yml', 'Main_Program.spec',
                 'AutoLychee.spec', 'Main_Program.py', 'I_Main.ico', 'setting.ico', 'Icon/Autolychee.png',
                 'updater.py', 'update_cache.py', 'release_files.py', 'release_build.py',
                 'All_Programs/158_AutoLychee_OneFile.py', 'All_Programs/one.py']
        for name in paths:
            path = self.root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text('import os\n' if name.endswith('.py') else 'fixture', encoding='utf-8')
        (self.root / 'Main_Program.py').write_text('CURRENT_VERSION = "1.1.97"\nimport os\n', encoding='utf-8')
        root_patch = patch.object(build, 'ROOT', self.root)
        root_patch.start()
        self.addCleanup(root_patch.stop)

    def change(self, name, text):
        (self.root / name).write_text(text, encoding='utf-8')

    def test_version_bump_does_not_rebuild_main(self):
        before = build.fingerprints()
        self.change('Main_Program.py', 'CURRENT_VERSION = "1.1.98"\nimport os\n')
        self.assertEqual(before, build.fingerprints())

    def test_one_program_same_imports_does_not_rebuild_runtimes(self):
        before = build.fingerprints()
        self.change('All_Programs/one.py', 'import os\nanswer = 42\n')
        self.assertEqual(before, build.fingerprints())

    def test_new_program_dependency_forces_main_build(self):
        before = build.fingerprints()
        self.change('All_Programs/one.py', 'import pandas\n')
        self.assertNotEqual(before['main'], build.fingerprints()['main'])

    def test_autolychee_change_rebuilds_only_autolychee(self):
        before = build.fingerprints()
        self.change('All_Programs/158_AutoLychee_OneFile.py', 'import os\nanswer = 43\n')
        after = build.fingerprints()
        self.assertEqual(before['main'], after['main'])
        self.assertEqual(before['updater'], after['updater'])
        self.assertNotEqual(before['autolychee'], after['autolychee'])

    def test_dependency_lock_change_rebuilds_all_components(self):
        before = build.fingerprints()
        self.change('requirements_ttk.lock.txt', 'new dependencies')
        after = build.fingerprints()
        self.assertTrue(all(before[key] != after[key] for key in before))

    def test_updater_change_rebuilds_only_updater(self):
        before = build.fingerprints()
        self.change('updater.py', 'import os\nanswer = 44\n')
        after = build.fingerprints()
        self.assertEqual(before['main'], after['main'])
        self.assertEqual(before['autolychee'], after['autolychee'])
        self.assertNotEqual(before['updater'], after['updater'])

    def test_reused_launcher_reads_release_version(self):
        actual_source = Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig')
        tree = ast.parse(actual_source)
        version_index = next(i for i, n in enumerate(tree.body) if isinstance(n, ast.Assign)
                             and any(isinstance(t, ast.Name) and t.id == 'CURRENT_VERSION' for t in n.targets))
        nodes = tree.body[version_index:version_index + 2]
        (self.root / 'release_version.json').write_text(json.dumps({'version': '1.1.98'}), encoding='utf-8')
        namespace = {'sys': SimpleNamespace(frozen=True, _MEIPASS=str(self.root), executable='launcher.exe'), 'os': os}
        exec(compile(ast.Module(body=nodes, type_ignores=[]), '<version metadata>', 'exec'), namespace)
        self.assertEqual(namespace['CURRENT_VERSION'], '1.1.98')

    def test_launcher_requests_recovery_before_loading_ui(self):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_recover_interrupted_update')
        executable = self.root / 'Main_Program.exe'
        (self.root / 'updater.exe').write_bytes(b'updater fixture')
        journal = self.root / '_internal/update-transactions/update-1/journal.json'
        journal.parent.mkdir(parents=True)
        journal.write_text('{"state":"applying"}', encoding='utf-8')
        namespace = {'sys': SimpleNamespace(frozen=True, executable=str(executable)), 'os': os, '_fast_show_error': Mock()}
        exec(compile(ast.Module(body=[function], type_ignores=[]), '<startup recovery>', 'exec'), namespace)
        with patch('subprocess.Popen') as launch:
            self.assertTrue(namespace['_recover_interrupted_update']())
        self.assertIn('--recover-only', launch.call_args.args[0])
        self.assertEqual(launch.call_args.kwargs['env']['PYINSTALLER_RESET_ENVIRONMENT'], '1')

    def test_missing_recovery_updater_prevents_partial_installation_start(self):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_recover_interrupted_update')
        journal = self.root / '_internal/update-transactions/update-1/journal.json'
        journal.parent.mkdir(parents=True)
        journal.write_text('{"state":"applying"}', encoding='utf-8')
        error = Mock()
        namespace = {'sys': SimpleNamespace(frozen=True, executable=str(self.root/'Main_Program.exe')), 'os': os, '_fast_show_error': error}
        exec(compile(ast.Module(body=[function], type_ignores=[]), '<startup recovery>', 'exec'), namespace)
        self.assertTrue(namespace['_recover_interrupted_update']())
        error.assert_called_once()

    def test_bridge_reconstructs_added_changed_deleted_and_combines_updates(self):
        old, middle, new = [self.root / name for name in ('old', 'middle', 'new')]
        for folder in (old, middle, new):
            folder.mkdir()
        for folder, values in ((old, {'changed': b'old', 'deleted': b'delete', 'same': b'same'}),
                               (middle, {'changed': b'middle', 'added': b'add', 'same': b'same'}),
                               (new, {'changed': b'new', 'same': b'same', 'deleted': b'readded'})):
            for name, value in values.items():
                (folder / name).write_bytes(value)
        first = self.root / 'first.zip'
        make_file_package(inventory(old), middle, '1.1.96', '1.1.99', first)
        with zipfile.ZipFile(first) as archive:
            manifest = json.loads(archive.read('manifest.json'))
        reconstructed = build.reconstruct_base_files(inventory(middle), manifest, '1.1.96', '1.1.99')
        self.assertEqual({k:v['sha256'] for k,v in reconstructed.items()},
                         {k:v['sha256'] for k,v in inventory(old).items()})
        combined = self.root / 'combined.zip'
        make_file_package(reconstructed, new, '1.1.96', '1.1.100', combined, inventory(new))
        apply_file_update(old, combined, '1.1.96', '1.1.100')
        self.assertEqual(inventory(old), inventory(new))
        manifest['files'][0]['after'] = '0' * 64
        with self.assertRaises(ValueError):
            build.reconstruct_base_files(inventory(middle), manifest, '1.1.96', '1.1.99')

    def updater_helper(self):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'ensure_updater_executable')
        namespace = {'os': os}
        exec(compile(ast.Module(body=[function], type_ignores=[]), '<updater download>', 'exec'), namespace)
        return namespace['ensure_updater_executable']

    def test_identical_updater_skips_download(self):
        target = self.root / 'updater.exe'
        target.write_bytes(b'MZexisting')
        checksum = 'sha256:' + hashlib.sha256(target.read_bytes()).hexdigest()
        with patch('urllib.request.urlopen') as network:
            self.assertFalse(self.updater_helper()('https://example.com/updater.exe', str(target), checksum))
            network.assert_not_called()

    def test_updater_download_is_verified_before_replacement(self):
        target = self.root / 'updater.exe'
        target.write_bytes(b'MZold')
        content = b'MZnew'
        checksum = 'sha256:' + hashlib.sha256(content).hexdigest()
        with patch('urllib.request.urlopen', return_value=io.BytesIO(content)):
            self.assertTrue(self.updater_helper()('https://example.com/updater.exe', str(target), checksum))
        self.assertEqual(target.read_bytes(), content)

    def test_invalid_or_interrupted_download_preserves_updater(self):
        target = self.root / 'updater.exe'
        target.write_bytes(b'MZold')
        for content in (b'MZwrong', b'not an exe'):
            with patch('urllib.request.urlopen', return_value=io.BytesIO(content)):
                with self.assertRaises(ValueError):
                    self.updater_helper()('https://example.com/updater.exe', str(target), 'sha256:' + '0'*64)
            self.assertEqual(target.read_bytes(), b'MZold')
        with patch('urllib.request.urlopen', side_effect=OSError('network interrupted')):
            with self.assertRaises(OSError):
                self.updater_helper()('https://example.com/updater.exe', str(target))
        self.assertEqual(target.read_bytes(), b'MZold')
        self.assertFalse(list(self.root.glob('*.download')))


if __name__ == '__main__':
    unittest.main()
