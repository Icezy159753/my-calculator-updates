import contextlib
import io
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from tools import deploy_release as deploy


class DeployTests(unittest.TestCase):
    def test_versions_keep_manual_bump_and_skip_used_tags(self):
        self.assertEqual(deploy.next_version('1.1.104', ['v1.1.104', 'v1.1.105']), '1.1.106')
        self.assertEqual(deploy.next_version('1.1.110', ['v1.1.104']), '1.1.110')
        self.assertEqual(deploy.next_version('1.1.104', ['not-a-release']), '1.1.104')

    def test_preflight_failure_stops_before_other_tests_or_smokes(self):
        with patch.object(deploy, 'run_preflight', side_effect=RuntimeError('missing CI dependency')), \
             patch.object(deploy, 'run') as runner, contextlib.redirect_stdout(io.StringIO()):
            with self.assertRaises(RuntimeError): deploy.release_gates()
        runner.assert_not_called()

    def test_only_release_inputs_are_added(self):
        for name in ('All_Programs/new.py', 'Main_Program.py', 'Icon/new.png', 'Deploy_GitHub.bat'):
            self.assertTrue(deploy.release_input(name))
        for name in ('Test3.json', 'openrouter.json', '.env', 'client-job.xlsx', 'outputs/raw.csv', 'build/output.exe'):
            self.assertFalse(deploy.release_input(name))

    def fixture(self, folder):
        root = Path(folder)
        (root / 'All_Programs').mkdir()
        (root / 'All_Programs/one.py').write_text('answer = 42\n', encoding='utf-8')
        main = root / 'Main_Program.py'
        main.write_text('CURRENT_VERSION = "1.1.104"\n', encoding='utf-8')
        instructions = root / 'Build exe Update Github.txt'
        instructions.write_text('v1.1.104\n', encoding='utf-8')
        def git(*args):
            if args == ('branch', '--show-current'): return 'main'
            if args == ('remote', 'get-url', 'origin'): return 'https://github.com/' + deploy.REPO + '.git'
            if args[0] == 'rev-list': return '0 0'
            if args == ('tag', '-l', 'v*'): return 'v1.1.104'
            if args == ('diff', '--name-only', '-z', 'HEAD'):
                values = ['All_Programs/one.py']
                if '1.1.105' in main.read_text(): values += ['Main_Program.py', instructions.name]
                return '\0'.join(values) + '\0'
            if args[0] == 'ls-files': return 'client-job.xlsx\0'
            return ''
        return root, main, instructions, Mock(side_effect=git)

    def exercise(self, action, answer=''):
        with TemporaryDirectory() as folder:
            root, main, instructions, git = self.fixture(folder)
            before = (main.read_bytes(), instructions.read_bytes())
            runner = Mock(return_value='')
            options = SimpleNamespace(dry_run=False, version_only=False, yes=False, no_watch=True)
            with patch.object(deploy, 'ROOT', root), patch.object(deploy, 'git', git), \
                 patch.object(deploy, 'run', runner), patch.object(deploy.shutil, 'which', return_value='gh'), \
                 patch.object(deploy, 'release_gates', side_effect=lambda: action(root)), \
                 patch('builtins.input', return_value=answer), contextlib.redirect_stdout(io.StringIO()):
                failure = None
                try: deploy.deploy(options)
                except RuntimeError as error: failure = error
            return before, (main.read_bytes(), instructions.read_bytes()), runner, failure

    def test_failed_checks_restore_version_and_never_commit_or_push(self):
        def fail(root): raise RuntimeError('failed smoke')
        before, after, runner, error = self.exercise(fail)
        self.assertEqual(before, after)
        self.assertIsNotNone(error)
        self.assertFalse(any(call.args[0][0] == 'git' for call in runner.call_args_list))

    def test_cancel_restores_version(self):
        before, after, runner, error = self.exercise(lambda root: None, 'n')
        self.assertEqual(before, after)
        self.assertIsNone(error)
        self.assertFalse(any(call.args[0][0] == 'git' for call in runner.call_args_list))

    def test_edit_during_checks_stops_publication(self):
        def edit(root): (root / 'All_Programs/one.py').write_text('answer = 99\n')
        before, after, runner, error = self.exercise(edit)
        self.assertIsNotNone(error)
        self.assertEqual(before, after)
        self.assertFalse(any(call.args[0][0] == 'git' for call in runner.call_args_list))

    def test_success_stages_selected_files_and_pushes_both_refs_atomically(self):
        before, after, runner, error = self.exercise(lambda root: None)
        self.assertIsNone(error)
        self.assertIn(b'1.1.105', after[0])
        commands = [call.args[0] for call in runner.call_args_list]
        self.assertIn(['git', 'push', '--atomic', 'origin', 'main', 'v1.1.105'], commands)
        staged = next(command for command in commands if command[:2] == ['git', 'add'])
        self.assertNotIn('client-job.xlsx', staged)

    def test_failed_push_retry_reuses_tag_without_new_commit_or_version(self):
        with TemporaryDirectory() as folder:
            root, main, instructions, fixture_git = self.fixture(folder)
            def git(*args):
                if args[0] == 'rev-list': return '1 0'
                if args[0] == 'rev-parse': return 'abc123'
                if args[0] in ('diff', 'ls-files', 'ls-remote'): return ''
                return fixture_git(*args)
            runner = Mock(return_value='')
            options = SimpleNamespace(dry_run=False, version_only=False, yes=True, no_watch=True)
            before = main.read_bytes()
            with patch.object(deploy, 'ROOT', root), patch.object(deploy, 'git', side_effect=git), \
                 patch.object(deploy, 'run', runner), patch.object(deploy, 'release_gates'), \
                 patch.object(deploy.shutil, 'which', return_value='gh'), contextlib.redirect_stdout(io.StringIO()):
                deploy.deploy(options)
            self.assertEqual(main.read_bytes(), before)
            commands = [call.args[0] for call in runner.call_args_list]
            self.assertIn(['git', 'push', '--atomic', 'origin', 'main', 'v1.1.104'], commands)
            self.assertFalse(any(command[:2] in (['git', 'add'], ['git', 'commit']) for command in commands))


if __name__ == '__main__':
    unittest.main()
