"""Regression checks for launcher arguments leaking into SPSS file loading."""
import ast
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock


MODULE = Path(__file__).parent / 'All_Programs' / '155_TRF SPSS Excel 01.py'


class LaunchArgumentsTest(unittest.TestCase):
    def run_entry(self, process_args, **kwargs):
        tree = ast.parse(MODULE.read_text(encoding='utf-8-sig'))
        entry = next(node for node in tree.body
                     if isinstance(node, ast.FunctionDef) and node.name == 'run_this_app')
        app = MagicMock()
        window = MagicMock()
        qt = MagicMock(return_value=app)
        namespace = dict(
            sys=SimpleNamespace(argv=process_args, platform='test', exit=MagicMock()),
            QApplication=qt, Window=MagicMock(return_value=window),
            APP_NAME='TRF', STYLE='', app_icon=lambda: None,
        )
        exec(compile(ast.Module(body=[entry], type_ignores=[]), str(MODULE), 'exec'), namespace)
        namespace['run_this_app'](**kwargs)
        app.exec.assert_called_once()
        return window, qt

    def test_launcher_flags_are_not_loaded_as_a_file(self):
        args = ['Main_Program.py', '--run-module', MODULE.stem,
                '--entry-point', 'run_this_app', '--working-dir', 'C:/Survey Job']
        window, qt = self.run_entry(args, working_dir='C:/Survey Job')
        window.load.assert_not_called()
        qt.assert_called_once_with(['Main_Program.py'])

    def test_direct_execution_without_file(self):
        window, _ = self.run_entry([str(MODULE)], argv=[])
        window.load.assert_not_called()

    def test_explicit_file_with_spaces(self):
        path = 'C:/Survey Job/raw data.sav'
        window, qt = self.run_entry([str(MODULE), path], argv=[path])
        window.load.assert_called_once_with(path)
        qt.assert_called_once_with([str(MODULE), path])


if __name__ == '__main__':
    unittest.main()
