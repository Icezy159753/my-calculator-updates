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


if __name__ == '__main__':
    unittest.main()
