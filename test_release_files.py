import json
import os
from pathlib import Path
import shutil
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
import zipfile

import release_files as files


class ReleaseFilesTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.old = self.root / 'old'
        self.new = self.root / 'new'
        self.app = self.root / 'app'
        for directory in (self.old, self.new, self.app):
            directory.mkdir()
        self.write(self.old, 'Main_Program.exe', b'old launcher')
        self.write(self.old, '_internal/All_Programs/program.py', b'old program')
        self.write(self.old, '_internal/obsolete.dll', b'old managed library')
        shutil.copytree(self.old, self.new, dirs_exist_ok=True)
        self.write(self.new, 'Main_Program.exe', b'new launcher')
        self.write(self.new, '_internal/All_Programs/program.py', b'new program')
        self.write(self.new, '_internal/new.dll', b'new library')
        (self.new / '_internal/obsolete.dll').unlink()
        self.write(self.new, '_internal/release_version.json', b'{"version":"1.1.97"}')
        shutil.copytree(self.old, self.app, dirs_exist_ok=True)
        self.package = self.root / 'update.zip'
        files.make_file_package(files.inventory(self.old), self.new, '1.1.96', '1.1.97', self.package)

    @staticmethod
    def write(root, name, content):
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)

    def apply(self):
        files.apply_file_update(self.app, self.package, '1.1.96', '1.1.97')

    def test_updates_adds_and_deletes_match_full_package(self):
        self.apply()
        self.assertEqual(files.inventory(self.app), files.inventory(self.new))

    def test_user_files_and_settings_are_preserved(self):
        for name in files.PROTECTED:
            self.write(self.app, name, b'user data')
        self.write(self.app, 'my-survey.xlsx', b'user survey')
        self.apply()
        for name in files.PROTECTED:
            self.assertEqual((self.app / name).read_bytes(), b'user data')
        self.assertEqual((self.app / 'my-survey.xlsx').read_bytes(), b'user survey')

    def test_wrong_base_rejected_before_any_file_changes(self):
        self.write(self.app, '_internal/All_Programs/program.py', b'local modified program')
        before = files.inventory(self.app)
        with self.assertRaises(files.UpdateRejected):
            self.apply()
        self.assertEqual(files.inventory(self.app), before)

    def test_skipped_version_rejected(self):
        with self.assertRaises(files.UpdateRejected):
            files.apply_file_update(self.app, self.package, '1.1.94', '1.1.97')
        self.assertEqual(files.inventory(self.app), files.inventory(self.old))

    def test_locked_file_rolls_back_every_change(self):
        replace = os.replace

        def locked_replace(source, target):
            if '/stage/' in str(source).replace('\\', '/') and str(target).endswith('program.py'):
                raise PermissionError('in use')
            return replace(source, target)

        with patch('release_files.os.replace', side_effect=locked_replace):
            with self.assertRaises(PermissionError):
                self.apply()
        self.assertEqual(files.inventory(self.app), files.inventory(self.old))

    def test_crash_after_first_replacement_recovers_on_next_run(self):
        replace = os.replace
        count = 0

        def crash(source, target):
            nonlocal count
            result = replace(source, target)
            if '/stage/' in str(source).replace('\\', '/'):
                count += 1
                if count == 1:
                    raise KeyboardInterrupt('simulated power interruption')
            return result

        with patch('release_files.os.replace', side_effect=crash):
            with self.assertRaises(KeyboardInterrupt):
                self.apply()
        files.recover_pending(self.app)
        self.assertEqual(files.inventory(self.app), files.inventory(self.old))

    def test_payload_corruption_rejected(self):
        replacement = self.root / 'corrupt.zip'
        with zipfile.ZipFile(self.package) as source, zipfile.ZipFile(replacement, 'w') as target:
            for name in source.namelist():
                data = source.read(name)
                if name == 'payload/Main_Program.exe':
                    data = b'x' * len(data)
                target.writestr(name, data)
        self.package = replacement
        with self.assertRaises(files.UpdateRejected):
            self.apply()
        self.assertEqual(files.inventory(self.app), files.inventory(self.old))

    def test_path_traversal_and_windows_special_paths_rejected(self):
        for name in ('../outside', '/outside', 'C:/outside', 'folder\\outside', 'CON.txt', 'x/../y', 'x./y'):
            with self.subTest(name=name), self.assertRaises(files.UpdateRejected):
                files.safe_name(name)

    def test_unexpected_zip_member_rejected(self):
        with zipfile.ZipFile(self.package, 'a') as archive:
            archive.writestr('../outside', b'bad')
        with self.assertRaises(files.UpdateRejected):
            self.apply()
        self.assertEqual(files.inventory(self.app), files.inventory(self.old))

    def test_cache_and_transactions_are_not_release_payloads(self):
        self.write(self.new, '_internal/updates/package_1.1.96.zip', b'cache')
        self.assertNotIn('_internal/updates/package_1.1.96.zip', files.inventory(self.new))

    def test_full_repair_preserves_user_data_and_replaces_corrupt_program(self):
        self.write(self.app, 'my-survey.xlsx', b'user work')
        self.write(self.app, '_internal/Test3.json', b'user settings')
        self.write(self.app, 'Main_Program.exe', b'corrupt')
        files.install_full_update(self.app, self.new, '1.1.94', '1.1.97', self.root)
        self.assertEqual((self.app / 'Main_Program.exe').read_bytes(), b'new launcher')
        self.assertEqual((self.app / 'my-survey.xlsx').read_bytes(), b'user work')
        self.assertEqual((self.app / '_internal/Test3.json').read_bytes(), b'user settings')

    def test_full_install_locked_file_restores_originals(self):
        before = files.inventory(self.app)
        replace = os.replace

        def locked_replace(source, target):
            if '/stage/' in str(source).replace('\\', '/') and str(target).endswith('program.py'):
                raise PermissionError('in use')
            return replace(source, target)

        with patch('release_files.os.replace', side_effect=locked_replace):
            with self.assertRaises(PermissionError):
                files.install_full_update(self.app, self.new, '1.1.96', '1.1.97', self.root)
        self.assertEqual(files.inventory(self.app), before)

    def test_failed_rollback_retains_backup_for_recovery(self):
        replace = os.replace

        def locked_replace(source, target):
            if '/stage/' in str(source).replace('\\', '/') and str(target).endswith('program.py'):
                raise PermissionError('in use')
            if str(source).endswith('.rollback-tmp'):
                raise PermissionError('rollback file in use')
            return replace(source, target)

        with patch('release_files.os.replace', side_effect=locked_replace):
            with self.assertRaises(files.RollbackFailed):
                self.apply()
        journals = list((self.app / '_internal/update-transactions').glob('*/journal.json'))
        self.assertEqual(len(journals), 1)
        files.recover_pending(self.app)
        self.assertEqual(files.inventory(self.app), files.inventory(self.old))

    def test_small_program_change_does_not_include_unchanged_runtime(self):
        runtime = os.urandom(512000)
        self.write(self.old, '_internal/runtime.dll', runtime)
        self.write(self.new, '_internal/runtime.dll', runtime)
        files.make_file_package(files.inventory(self.old), self.new, '1.1.96', '1.1.97', self.package)
        self.assertLess(self.package.stat().st_size, len(runtime) // 20)

    def test_full_zip_traversal_rejected_before_extraction(self):
        package = self.root / 'bad-full.zip'
        with zipfile.ZipFile(package, 'w') as archive:
            archive.writestr('Main_Program.exe', b'valid-looking first entry')
            archive.writestr('../outside', b'bad')
        with self.assertRaises(files.UpdateRejected):
            files.extract_full_package(package, self.root / 'extracted')
        self.assertFalse((self.root / 'extracted/Main_Program.exe').exists())


if __name__ == '__main__':
    unittest.main()
