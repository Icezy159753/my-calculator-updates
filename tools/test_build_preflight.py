from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from tools import build_preflight as preflight


class PreflightTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for name in ('Main_Program.py', 'updater.py', 'release_build.py', 'release_files.py', 'update_cache.py',
                     'verify_file_transition.py', 'All_Programs/158_AutoLychee_OneFile.py', 'All_Programs/one.py'):
            self.write(name, 'import os\n')
        self.write('Main_Program.spec', "hiddenimports = ['known']\ndatas=[('template.xlsx', '.') ]\na=Analysis([],excludes=['PySide6'])\nexe=EXE(icon='I_Main.ico')\n")
        self.write('AutoLychee.spec', "hiddenimports=['PySide6.QtWidgets']\na=Analysis([],excludes=['PyQt6'])\n")
        self.write('requirements_ttk.lock.txt', 'pandas==1.0.0\nrequests==2.0.0\nknown==1.0.0\n')
        self.write('.github/workflows/windows-release.yml', '''if ("${{ steps.plan.outputs.main }}" -eq "true") {
 python -m pip install -r requirements_ttk.lock.txt
}
if ("${{ steps.plan.outputs.autolychee }}" -eq "true") {
 python -m pip install PySide6==6.9.0
}
if ("${{ steps.plan.outputs.updater }}" -eq "true") {
 python -m pip install requests==2.0.0
}
''')
        self.write('template.xlsx', 'fixture')
        self.write('I_Main.ico', 'fixture')
        self.baseline = {p.relative_to(self.root).as_posix(): p.read_text() for p in self.root.rglob('*.py')}

    def write(self, name, source):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(source, encoding='utf-8')

    def check(self):
        return preflight.check(self.root, self.baseline, mapping={})

    def test_current_sources_pass_and_scanner_does_not_execute_programs(self):
        self.write('All_Programs/one.py', "import os\nraise RuntimeError('do not execute me')\n")
        errors, warnings, count = self.check()
        self.assertEqual(errors, [])
        self.assertGreater(count, 0)

    def test_syntax_and_invalid_scope_are_blocked_with_file_and_line(self):
        for source in ('def broken(:\n', 'return 1\n'):
            self.write('All_Programs/one.py', source)
            errors, warnings, count = self.check()
            self.assertTrue(any('All_Programs/one.py:1:' in error for error in errors))

    def test_new_required_package_not_in_ci_is_blocked(self):
        self.write('All_Programs/new.py', 'import definitely_missing_package\n')
        errors, warnings, count = self.check()
        self.assertTrue(any('definitely_missing_package' in error and 'CI' in error for error in errors))

    def test_optional_and_type_checking_imports_are_not_hard_failures(self):
        self.write('All_Programs/one.py', 'try:\n import optional_package\nexcept ImportError:\n pass\nfrom typing import TYPE_CHECKING\nif TYPE_CHECKING:\n import types_only_package\n')
        errors, warnings, count = self.check()
        self.assertEqual(errors, [])
        self.assertTrue(any('optional_package' in warning for warning in warnings))
        self.assertFalse(any('types_only_package' in warning for warning in warnings))

    def test_declared_new_package_warns_about_hidden_import_until_spec_added(self):
        self.write('All_Programs/new.py', 'import pandas\n')
        errors, warnings, count = self.check()
        self.assertEqual(errors, [])
        self.assertTrue(any('hiddenimports' in warning for warning in warnings))
        with (self.root / 'Main_Program.spec').open('a') as stream: stream.write("\nhiddenimports += ['pandas']\n")
        errors, warnings, count = self.check()
        self.assertEqual(errors, [])
        self.assertEqual(warnings, [])

    def test_qt_exclusion_and_relative_module_deletion_are_blocked(self):
        self.write('All_Programs/new.py', 'from PySide6.QtWidgets import QWidget\n')
        errors, warnings, count = self.check()
        self.assertTrue(any('excludes' in error for error in errors))
        self.write('All_Programs/new.py', 'from .gone import helper\n')
        errors, warnings, count = self.check()
        self.assertTrue(any('relative' in error for error in errors))
        self.baseline['All_Programs/gone.py'] = 'helper = 1\n'
        self.baseline['All_Programs/new.py'] = 'from .gone import helper\n'
        errors, warnings, count = self.check()
        self.assertTrue(any('ถูกลบ' in error for error in errors))

    def test_missing_assets_are_blocked_and_dynamic_imports_warn(self):
        (self.root / 'I_Main.ico').unlink()
        self.write('All_Programs/one.py', 'import importlib\nmodule_name="plugin"\nimportlib.import_module(module_name)\n')
        errors, warnings, count = self.check()
        self.assertTrue(any('I_Main.ico' in error for error in errors))
        self.assertTrue(any('hiddenimports' in warning for warning in warnings))

    def test_packages_are_scoped_to_correct_ci_component(self):
        packages = preflight.ci_packages(self.root)
        self.assertIn('pandas', packages['main'])
        self.assertNotIn('pandas', packages['autolychee'])
        self.assertIn('pyside6', packages['autolychee'])

    def test_spec_imports_and_analysis_source_are_checked(self):
        self.baseline['Main_Program.spec'] = (self.root / 'Main_Program.spec').read_text()
        with (self.root / 'Main_Program.spec').open('a') as stream:
            stream.write("\nimport missing_build_library\nsource='missing_source.py'\nb=Analysis([source])\n")
        errors, warnings, count = self.check()
        self.assertTrue(any('missing_build_library' in error for error in errors))
        self.assertTrue(any('missing_source.py' in error for error in errors))

    def test_syntax_newer_than_ci_python_is_blocked(self):
        self.write('All_Programs/one.py', 'value = t"new syntax"\n')
        errors, warnings, count = self.check()
        self.assertTrue(any('Python 3.12' in error for error in errors))

    def test_new_root_helper_is_syntax_checked_and_bundle_warning_is_given(self):
        self.write('new_helper.py', 'def bad(:\n')
        self.write('All_Programs/new.py', 'import new_helper\n')
        errors, warnings, count = self.check()
        self.assertTrue(any('new_helper.py:1:' in error for error in errors))
        self.assertTrue(any('bundle source' in warning for warning in warnings))


if __name__ == '__main__':
    unittest.main()
