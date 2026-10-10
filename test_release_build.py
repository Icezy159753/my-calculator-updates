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
import threading
from release_files import inventory, make_file_package, apply_file_update

import release_build as build
import verify_file_transition as transition


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

    def test_transition_checks_section_count_from_current_source(self):
        work = self.root / 'build'
        dist = self.root / 'dist'
        work.mkdir()
        dist.mkdir()
        (work / 'plan.json').write_text(json.dumps({'previous': '1.1.109', 'version': '1.1.111'}))
        (dist / 'release_manifest.json').write_text(json.dumps({'files': {}}))
        for source in (self.root / 'All_Programs').glob('*.py'):
            bundled = dist / 'Main_Program/_internal/All_Programs' / source.name
            bundled.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, bundled)
        for count in (12, 13):
            for output, passes in ((f'{count} sections OK'.encode(), True),
                                   (f'{count - 1} sections OK'.encode(), False), (b'', False)):
                with self.subTest(count=count, output=output):
                    self.change('All_Programs/158_AutoLychee_OneFile.py', ''.join(
                        f'# ====== MODULE: section_{i} ======\n' for i in range(count)))
                    shutil.copy2(self.root / 'All_Programs/158_AutoLychee_OneFile.py',
                                 dist / 'Main_Program/_internal/All_Programs/158_AutoLychee_OneFile.py')
                    processes = []
                    for stdout in (b'Updater transaction self-test OK',
                                   b'Auto Lychee GUI smoke test OK', b'', output):
                        process = Mock(returncode=0)
                        process.communicate.return_value = (stdout, b'')
                        processes.append(process)
                    with patch.object(transition, 'ROOT', self.root), \
                         patch.object(transition, 'BUILD', work), patch.object(transition, 'DIST', dist), \
                         patch.object(transition, 'apply_file_update'), \
                         patch.object(transition, 'inventory', return_value={}), \
                         patch.object(transition.subprocess, 'Popen', side_effect=processes):
                        if passes:
                            transition.verify()
                        else:
                            with self.assertRaisesRegex(RuntimeError, f'{count} sections OK'):
                                transition.verify()

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

    def test_thai_console_diagnostics_do_not_fail_on_unicode(self):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_fast_write_output')
        namespace = {}
        exec(compile(ast.Module(body=[function], type_ignores=[]), '<console diagnostics>', 'exec'), namespace)
        buffer = io.BytesIO()
        stream = io.TextIOWrapper(buffer, encoding='cp874')
        namespace['_fast_write_output'](stream, '12 sections OK · smoke test')
        self.assertIn(b'12 sections OK', buffer.getvalue())
        namespace['_fast_write_output'](None, 'windowed app has no stream')

    def test_redirected_launcher_diagnostics_are_utf8_on_western_windows(self):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_fast_write_output')
        namespace = {}
        exec(compile(ast.Module(body=[function], type_ignores=[]), '<console diagnostics>', 'exec'), namespace)
        buffer = io.BytesIO()
        stream = io.TextIOWrapper(buffer, encoding='cp1252')
        text = '12 sections OK · data: แบบสอบถาม'
        namespace['_fast_write_output'](stream, text)
        self.assertEqual(buffer.getvalue().decode('utf-8'), text)

    def test_failed_auto_check_exits_without_error_dialog(self):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_fast_launch_submodule')
        error = Mock()
        namespace = {'os': os, 'sys': SimpleNamespace(frozen=True, executable=str(self.root/'Main_Program.exe'),
                     _MEIPASS=str(self.root), path=list(sys.path), argv=['Main_Program.exe','--run-module','158_AutoLychee_OneFile','--check']),
                     '_fast_show_error': error}
        exec(compile(ast.Module(body=[function], type_ignores=[]), '<fast check>', 'exec'), namespace)
        with patch('traceback.print_exc'), self.assertRaises(SystemExit) as result:
            namespace['_fast_launch_submodule']()
        self.assertEqual(result.exception.code, 1)
        error.assert_not_called()

    def test_updater_uses_same_notification_credentials_as_main(self):
        configs = []
        for name in ('Main_Program.py', 'updater.py'):
            tree = ast.parse(Path(__file__).with_name(name).read_text(encoding='utf-8-sig'))
            values = {target.id: ast.literal_eval(node.value) for node in tree.body if isinstance(node, ast.Assign)
                      for target in node.targets if isinstance(target, ast.Name)
                      and target.id in ('TELEGRAM_BOT_TOKEN', 'TELEGRAM_CHAT_ID')}
            configs.append(values)
        # Do not expose credential values if this assertion fails.
        self.assertTrue(configs[0] == configs[1], 'Main/updater notification credentials differ')

    def notice_fixture(self, response=None, error=None):
        tree = ast.parse(Path(__file__).with_name('updater.py').read_text(encoding='utf-8-sig'))
        app = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'UpdaterApp')
        method = next(n for n in app.body if isinstance(n, ast.FunctionDef) and n.name == '_send_telegram_update_notice')
        send = Mock(return_value=response, side_effect=error)
        log = Mock()
        namespace = {'TELEGRAM_BOT_TOKEN': 'fixture-token', 'TELEGRAM_CHAT_ID': 'fixture-chat',
                     'requests': SimpleNamespace(post=send), '_log_update_event': log}
        exec(compile(ast.Module(body=[method], type_ignores=[]), '<update notice>', 'exec'), namespace)
        instance = SimpleNamespace(current_version='1.1.96', new_version='1.1.102',
                                   release_url='https://example.com/release',
                                   _get_user_machine_info=lambda: ('user<&>', 'PC&one', '192.168.1.42'))
        return namespace['_send_telegram_update_notice'], instance, send, log

    def test_success_notice_contains_machine_versions_and_release(self):
        method, instance, send, log = self.notice_fixture(SimpleNamespace(status_code=200, json=lambda: {'ok': True}))
        self.assertTrue(method(instance))
        payload = send.call_args.kwargs['json']
        for text in ('user&lt;&amp;&gt;', 'PC&amp;one', '192.168.1.42', '1.1.96', '1.1.102', 'https://example.com/release'):
            self.assertIn(text, payload['text'])
        self.assertFalse(payload['disable_web_page_preview'])
        self.assertEqual(payload['parse_mode'], 'HTML')
        self.assertIn('sent=True', log.call_args.args[0])

    def test_notification_failure_is_logged_without_secrets_and_does_not_fail_install(self):
        method, instance, send, log = self.notice_fixture(SimpleNamespace(status_code=401, json=lambda: {'ok': False}))
        self.assertFalse(method(instance))
        self.assertIn('status=401', log.call_args.args[0])
        method, instance, send, log = self.notice_fixture(error=OSError('url contains fixture-token'))
        self.assertFalse(method(instance))
        self.assertNotIn('fixture-token', log.call_args.args[0])

    def startup_namespace(self, methods):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        launcher = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'AppLauncher')
        nodes = [n for n in launcher.body if isinstance(n, ast.FunctionDef) and n.name in methods]
        namespace = {'QtCore': SimpleNamespace(pyqtSlot=lambda *args: lambda method: method),
                     'CURRENT_VERSION': '1.1.102', '_normalize_tag_version': lambda tag: tag.lstrip('v')}
        exec(compile(ast.Module(body=nodes, type_ignores=[]), '<startup update>', 'exec'), namespace)
        return namespace

    def test_startup_new_release_prompts_and_current_release_does_not(self):
        ns = self.startup_namespace({'on_startup_release_checked'})
        window = SimpleNamespace(set_update_available=Mock(), set_update_status_latest=Mock(),
                                 set_update_status_error=Mock(), start_update_from_status_bar=Mock())
        versions = SimpleNamespace(parse=lambda value: tuple(map(int, value.split('.'))))
        with patch.dict(sys.modules, {'packaging.version': versions}):
            ns['on_startup_release_checked'](window, {'tag_name': 'v1.1.102'})
            window.set_update_status_latest.assert_called_once()
            window.start_update_from_status_bar.assert_not_called()
            release = {'tag_name': 'v1.1.103'}
            ns['on_startup_release_checked'](window, release)
            window.set_update_available.assert_called_once_with('1.1.103')
            window.start_update_from_status_bar.assert_called_once()
            self.assertIs(window._cached_update_release, release)

    def test_startup_worker_fetches_outside_ui_and_reports_failure(self):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        worker = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'StartupReleaseCheck')
        signals = []
        def signal(*args):
            value = SimpleNamespace(emit=Mock()); signals.append(value); return value
        namespace = {'QtCore': SimpleNamespace(QObject=object, pyqtSignal=signal), 'threading': threading,
                     'REPO_OWNER': 'owner', 'REPO_NAME': 'repo'}
        exec(compile(ast.Module(body=[worker], type_ignores=[]), '<release fetch>', 'exec'), namespace)
        request = Mock(return_value=SimpleNamespace(raise_for_status=lambda: None, json=lambda: {'tag_name': 'v1.1.103'}))
        with patch('threading.Thread') as thread, patch.dict(sys.modules, {'requests': SimpleNamespace(get=request)}):
            namespace['StartupReleaseCheck']().start()
            request.assert_not_called()
            self.assertTrue(thread.call_args.kwargs['daemon'])
            thread.return_value.start.assert_called_once()
            fetch = thread.call_args.kwargs['target']
            fetch()
            signals[0].emit.assert_called_once_with({'tag_name': 'v1.1.103'})
            request.side_effect = OSError('offline')
            fetch()
            signals[1].emit.assert_called_once()

    def test_cached_startup_release_does_not_repeat_network_request_for_prompt(self):
        tree = ast.parse(Path(__file__).with_name('Main_Program.py').read_text(encoding='utf-8-sig'))
        method = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'check_for_updates')
        ask = Mock(return_value=False)
        namespace = {'CURRENT_VERSION': '1.1.102', '_normalize_tag_version': lambda tag: tag.lstrip('v'), 'ask_yes_no': ask}
        exec(compile(ast.Module(body=[method], type_ignores=[]), '<cached prompt>', 'exec'), namespace)
        request = Mock()
        versions = SimpleNamespace(parse=lambda value: tuple(map(int, value.split('.'))))
        with patch.dict(sys.modules, {'requests': SimpleNamespace(get=request), 'packaging.version': versions}):
            namespace['check_for_updates'](object(), latest_release={'tag_name': 'v1.1.103'})
        ask.assert_called_once()
        request.assert_not_called()


if __name__ == '__main__':
    unittest.main()
