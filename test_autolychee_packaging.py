"""Launcher routing checks without loading the Launcher UI or starting Lychee jobs."""
import ast
import os
import runpy
from pathlib import Path
from types import ModuleType, SimpleNamespace
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

    def test_program_configuration_routes_autolychee_as_a_script(self):
        tree = ast.parse((ROOT / 'Main_Program.py').read_text(encoding='utf-8-sig'))
        programs = next(ast.literal_eval(node.value) for node in tree.body
                        if isinstance(node, ast.Assign) and any(isinstance(target, ast.Name) and target.id == 'PROGRAMS' for target in node.targets))
        info = next(program for program in programs if program.get('module_path') == '158_AutoLychee_OneFile')
        self.assertEqual(info['entry_point'], '__main__')
        self.assertEqual(info['frozen_executable'], 'AutoLychee/AutoLychee.exe')

    def test_fast_path_ignores_old_entry_point_for_autolychee(self):
        tree = ast.parse((ROOT / 'Main_Program.py').read_text(encoding='utf-8-sig'))
        method = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == '_fast_launch_submodule')
        namespace = {'os': os, 'sys': sys, '__file__': str(ROOT / 'Main_Program.py'), '_fast_show_error': Mock()}
        writer = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == '_fast_write_output')
        exec(compile(ast.Module(body=[writer, method], type_ignores=[]), '<fast-route>', 'exec'), namespace)
        with patch.object(sys, 'frozen', False, create=True), patch.object(sys, 'argv', ['Main_Program.py', '--run-module', '158_AutoLychee_OneFile', '--entry-point', 'run_this_app', '--check']), patch('runpy.run_path') as run:
            self.assertTrue(namespace['_fast_launch_submodule']())
            self.assertEqual(sys.argv[-1], '--check')
            run.assert_called_once_with(str(ROOT / 'All_Programs/158_AutoLychee_OneFile.py'), run_name='__main__')
        namespace['_fast_show_error'].assert_not_called()

    def test_autolychee_keeps_ci_smoke_test_flag(self):
        # CI runs AutoLychee.exe --smoke-test; without this branch the GUI starts and never exits.
        source = (ROOT / 'All_Programs/158_AutoLychee_OneFile.py').read_text(encoding='utf-8')
        self.assertIn("elif args == ['--smoke-test']:", source)
        self.assertIn("print('Auto Lychee GUI smoke test OK', flush=True)", source)

    def test_autolychee_build_keeps_runtime_outside_executable(self):
        # Qt must live beside the EXE so launch does not unpack it into TEMP each time.
        hooks = ModuleType('PyInstaller.utils.hooks')
        hooks.collect_submodules = lambda name: []
        analysis = SimpleNamespace(pure=[], scripts=['loader'], binaries=['Qt runtime'], datas=['source'])
        executable, collect = Mock(return_value='exe'), Mock()
        with patch.dict(sys.modules, {'PyInstaller.utils.hooks': hooks}):
            runpy.run_path(str(ROOT / 'AutoLychee.spec'), init_globals={
                'Analysis': Mock(return_value=analysis), 'PYZ': Mock(return_value='pyz'),
                'EXE': executable, 'COLLECT': collect,
            })
        self.assertTrue(executable.call_args.kwargs['exclude_binaries'])
        self.assertNotIn(analysis.binaries, executable.call_args.args)
        self.assertEqual(collect.call_args.args, ('exe', analysis.binaries, analysis.datas))
        self.assertEqual(collect.call_args.kwargs['name'], 'AutoLychee')

    def test_main_build_includes_whole_autolychee_bundle(self):
        tree = ast.parse((ROOT / 'Main_Program.spec').read_text(encoding='utf-8'))
        datas = next(node.value for node in tree.body if isinstance(node, ast.Assign)
                     and any(isinstance(target, ast.Name) and target.id == 'datas' for target in node.targets))
        namespace = {'autolychee_bundle': 'dist/AutoLychee'}
        bundled = eval(compile(ast.Expression(datas.elts[0]), '<bundle>', 'eval'), namespace)
        self.assertEqual(bundled, ('dist/AutoLychee', 'AutoLychee'))

    def test_frozen_fast_path_opens_separate_exe(self):
        tree = ast.parse((ROOT / 'Main_Program.py').read_text(encoding='utf-8-sig'))
        method = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == '_fast_launch_submodule')
        namespace = {'os': os, 'sys': sys, '__file__': str(ROOT / 'Main_Program.py'), '_fast_show_error': Mock()}
        writer = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == '_fast_write_output')
        exec(compile(ast.Module(body=[writer, method], type_ignores=[]), '<fast-route>', 'exec'), namespace)
        with patch.object(sys, 'frozen', True, create=True), patch.object(sys, '_MEIPASS', str(ROOT / 'dist'), create=True), patch.object(sys, 'argv', ['Main_Program.exe', '--run-module', '158_AutoLychee_OneFile', '--entry-point', 'run_this_app', '--check']), patch('os.path.isfile', return_value=True), patch('subprocess.run', return_value=SimpleNamespace(returncode=0, stdout='', stderr='')) as call:
            with self.assertRaises(SystemExit) as result:
                namespace['_fast_launch_submodule']()
            self.assertEqual(result.exception.code, 0)
            self.assertEqual(call.call_args.args[0], [str(ROOT / 'dist/AutoLychee/AutoLychee.exe'), '--check'])
            self.assertTrue(call.call_args.kwargs['capture_output'])


if __name__ == '__main__':
    unittest.main()
