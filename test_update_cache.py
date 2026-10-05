import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
import zipfile

from update_cache import find_cached_package, store_cached_package


class UpdateCacheTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.app = self.root / 'installed'
        self.app.mkdir()
        environment = patch.dict(os.environ, {'LOCALAPPDATA': str(self.root / 'user')})
        environment.start()
        self.addCleanup(environment.stop)
        self.package = self.root / 'download.zip'
        self.make_zip(self.package, 'new')

    @staticmethod
    def make_zip(path, text):
        path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path, 'w') as archive:
            archive.writestr('Main_Program.exe', text)

    def test_cache_is_outside_installation_and_can_be_found(self):
        saved = store_cached_package(self.app, '1.1.97', self.package)
        self.assertTrue(Path(saved).is_relative_to(self.root / 'user'))
        self.assertEqual(find_cached_package(self.app, '1.1.97'), saved)
        self.assertEqual(Path(saved).read_bytes(), self.package.read_bytes())

    def test_legacy_cache_still_works(self):
        legacy = self.app / '_internal/updates/package_1.1.96.zip'
        self.make_zip(legacy, 'old')
        self.assertEqual(find_cached_package(self.app, '1.1.96'), str(legacy))

    def test_corrupt_cache_is_not_selected(self):
        legacy = self.app / '_internal/updates/package_1.1.96.zip'
        legacy.parent.mkdir(parents=True)
        legacy.write_bytes(b'incomplete download')
        self.assertIsNone(find_cached_package(self.app, '1.1.96'))

    def test_failed_copy_preserves_previous_package(self):
        saved = Path(store_cached_package(self.app, '1.1.97', self.package))
        previous = saved.read_bytes()
        self.make_zip(self.package, 'replacement')
        with patch('update_cache.shutil.copyfile', side_effect=PermissionError('write denied')):
            with self.assertRaises(OSError):
                store_cached_package(self.app, '1.1.97', self.package)
        self.assertEqual(saved.read_bytes(), previous)
        self.assertEqual(list(saved.parent.glob('*.tmp')), [])

    def test_locked_user_cache_falls_back_to_legacy_location(self):
        replace = os.replace

        def replace_unless_user_cache(source, target):
            if Path(target).is_relative_to(self.root / 'user'):
                raise PermissionError('locked')
            return replace(source, target)

        with patch('update_cache.os.replace', side_effect=replace_unless_user_cache):
            saved = store_cached_package(self.app, '1.1.97', self.package)
        self.assertEqual(saved, str(self.app / '_internal/updates/package_1.1.97.zip'))
        self.assertEqual(find_cached_package(self.app, '1.1.97'), saved)
        self.assertEqual(Path(saved).read_bytes(), self.package.read_bytes())

    def test_bad_source_cannot_overwrite_valid_cache(self):
        saved = Path(store_cached_package(self.app, '1.1.97', self.package))
        previous = saved.read_bytes()
        self.package.write_bytes(b'broken')
        with self.assertRaises(ValueError):
            store_cached_package(self.app, '1.1.97', self.package)
        self.assertEqual(saved.read_bytes(), previous)

    def test_installations_do_not_share_the_same_cache(self):
        first = store_cached_package(self.app, '1.1.97', self.package)
        second = store_cached_package(self.root / 'other', '1.1.97', self.package)
        self.assertNotEqual(first, second)


if __name__ == '__main__':
    unittest.main()
