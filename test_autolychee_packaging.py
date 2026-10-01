"""Launcher routing checks without loading the Launcher UI or starting Lychee jobs."""
import ast
import os
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
import sys

ROOT = Path(__file__).resolve().parent


class AutoLycheePackagingTests(unittest.TestCase):
    def setUp(self):
        source = (ROOT / 'Main_Program.py').read_text(encoding='utf-8-sig')
        tree = ast.parse(source)
        launcher = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == 'AppLauncher')
        method = next(node for node in launcher.body if isinstance(node, ast.FunctionDef) and node.name == '_start_local_module_process')
        self.popen = Mock(return_value=Mock())
        timer = Mock()
        thread = Mock()
        thread.side_effect = lambda target, **kwargs: SimpleNamespace(start=target)
        namespace = {
            'os': os, 'sys': sys, '__file__': str(ROOT / 'Main_Program.py'),
            'resource_path': lambda relative: str(ROOT / 'dist' / relative),
            'subprocess': SimpleNamespace(Popen=self.popen),
            'threading': SimpleNamespace(Thread=thread),
            'QtCore': SimpleNamespace(QTimer=Mock(return_value=timer)),
        }
        exec(compile(ast.Module(body=[method], type_ignores=[]), '<launcher-route>', 'exec'), namespace)
        self.start = namespace['_start_local_module_process']
        self.launcher = SimpleNamespace(launcher_base_dir=str(ROOT), _check_popen_ready=Mock())
        self.info = {'frozen_executable': 'AutoLychee/AutoLychee.exe'}

    def test_frozen_launches_bundled_exe_with_fresh_environment(self):
        with patch.object(sys, 'frozen', True, create=True), patch('os.path.isfile', return_value=True), patch.dict(os.environ, {'MAIN_PROGRAM_SCRIPT_MODULE': 'old-module'}):
            self.start(self.launcher, '158_AutoLychee_OneFile', '__main__', {}, self.info)
        args, kwargs = self.popen.call_args
        self.assertEqual(args[0], [str(ROOT / 'dist/AutoLychee/AutoLychee.exe')])
        self.assertEqual(kwargs['env']['PYINSTALLER_RESET_ENVIRONMENT'], '1')
        self.assertNotIn('MAIN_PROGRAM_SCRIPT_MODULE', kwargs['env'])

    def test_missing_bundled_exe_reports_error_without_import_fallback(self):
        with patch.object(sys, 'frozen', True, create=True), patch('os.path.isfile', return_value=False):
            self.start(self.launcher, '158_AutoLychee_OneFile', '__main__', {}, self.info)
        self.popen.assert_not_called()
        self.assertIsNotNone(self.launcher._pending_popen_error)

    def test_source_still_uses_script_loader(self):
        with patch.object(sys, 'frozen', False, create=True):
            self.start(self.launcher, '158_AutoLychee_OneFile', '__main__', {'working_dir': str(ROOT)}, self.info)
        args = self.popen.call_args.args[0]
        self.assertEqual(args[:3], [sys.executable, '-X', 'utf8'])
        self.assertIn('--run-module', args)
        self.assertEqual(args[-2:], ['--working-dir', str(ROOT)])

    def test_other_frozen_modules_keep_existing_entry_point(self):
        with patch.object(sys, 'frozen', True, create=True):
            self.start(self.launcher, 'existing_tool', 'run_this_app', {}, {})
        self.assertEqual(self.popen.call_args.args[0], [sys.executable, '--run-module', 'existing_tool', '--entry-point', 'run_this_app'])


if __name__ == '__main__':
    unittest.main()
