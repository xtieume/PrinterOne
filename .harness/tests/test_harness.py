"""Self-test for scripts/harness (5 commands) and the guard hook. Requirement ids: docs/specs/2026-11-07-harness.md.

Run: python3 -m unittest discover -s .harness/tests
Each case seeds one fault in a throwaway git repo with a tiny Python project. goalrun comes from the installed
testcase plugin (or GOALRUN=<path>, as CI sets it); it is never copied into this repo.
"""
import importlib.machinery
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]


def load(name, path):
    loader = importlib.machinery.SourceFileLoader(name, str(path))
    spec = importlib.util.spec_from_loader(name, loader)
    mod = importlib.util.module_from_spec(spec)
    loader.exec_module(mod)
    return mod


H = load('harness_facade', REPO / 'scripts' / 'harness')
G = load('harness_guard', REPO / '.harness' / 'hooks' / 'guard.py')
GOALRUN = str(H.goalrun_path(json.loads((REPO / '.harness/profile.json').read_text())))

RUNNER = '''import sys, unittest
args = sys.argv[1:]
pattern = args[args.index('--tests') + 1] if '--tests' in args else None
argv = ['run', 'discover', '-s', 'tests', '-t', '.'] + (['-k', pattern] if pattern else [])
unittest.main(module=None, argv=argv)
'''
CALC = 'def add(a, b):\n    return a + b\n'
TEST_ADD = ('import unittest\nfrom src import calc\n\n\nclass Add(unittest.TestCase):\n'
            '    def test_add(self):\n        self.assertEqual(calc.add(2, 3), 5)\n')
TEST_MUL = ('import unittest\nfrom src import calc\n\n\nclass Mul(unittest.TestCase):\n'
            '    def test_mul(self):\n        self.assertEqual(calc.mul(2, 3), 6)\n')
LEDGER_HEAD = '# id\twhat\tcheck\tdeliverable\tbreak\n'
MUL_ROW = ('MUL\tREQ-CALC-001 mul multiplies\tscripts/harness check unit --tests test_mul\tsrc/calc.py\t'
           'src/calc.py :: return a * b :: return a + b\n')


def git(cwd, *args):
    return subprocess.run(['git', *args], cwd=cwd, check=True, capture_output=True, text=True).stdout.strip()


def init_repo(path):
    """git init on main with a local identity: CI runners have no global one."""
    git(path, 'init', '-q', '-b', 'main')
    git(path, 'config', 'user.name', 't')
    git(path, 'config', 'user.email', 't@t')
    return path


class Fixture(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template = Path(tempfile.mkdtemp(prefix='harness-tpl-'))
        root = cls.template / 'repo'
        (root / 'scripts').mkdir(parents=True)
        shutil.copy2(REPO / 'scripts' / 'harness', root / 'scripts' / 'harness')
        real = json.loads((REPO / '.harness' / 'profile.json').read_text())
        profile = {
            'schema': 1, 'base_branch': 'main', 'approvers': ['owner'], 'goalrun': real['goalrun'],
            'checks': {
                'unit': {'argv': ['python3', 'tests/run.py'], 'kind': 'unittest', 'needs': ['python3'],
                         'paths': ['src/*', 'tests/*'], 'tests_filter': '^[A-Za-z0-9_.*]+$', 'timeout': 120},
                'echo-secret': {'argv': ['sh', '-c', 'echo "env=$FIXTURE_TOKEN pat=ghp_' + 'a' * 36 + '"'],
                                'kind': 'exit', 'needs': ['sh'], 'paths': []},
                'needs-missing': {'argv': ['true'], 'kind': 'exit', 'needs': ['no-such-tool-xyz'], 'paths': []},
                'docs': {'argv': ['true'], 'kind': 'exit', 'needs': [], 'paths': ['docs/*']},
                'one-file': {'argv': ['true'], 'kind': 'exit', 'needs': [], 'paths': [], 'tests_required': True,
                             'tests_filter': '^[a-z]+$'},
            },
            'product_paths': ['src/*'],
        }
        feats = {'features': [
            {'id': 'F-001', 'revision': 1, 'behavior': 'calc can multiply', 'spec': 'docs/spec.md',
             'acceptance': ['AC-1'], 'allowed_paths': ['src/*', 'tests/*'], 'required_checks': ['unit'],
             'risk': 'low', 'dependencies': []},
        ]}
        files = {
            '.harness/profile.json': json.dumps(profile, indent=2),
            '.harness/features.json': json.dumps(feats, indent=2),
            'docs/spec.md': '# Calc\nREQ-CALC-001 mul multiplies two numbers.\n',
            'src/__init__.py': '', 'src/calc.py': CALC,
            'tests/__init__.py': '', 'tests/run.py': RUNNER, 'tests/test_add.py': TEST_ADD,
            '.gitignore': '.testcases/\n__pycache__/\n',
        }
        for name, text in files.items():
            (root / name).parent.mkdir(parents=True, exist_ok=True)
            (root / name).write_text(text)
        init_repo(root)
        git(root, 'add', '-A')
        git(root, 'commit', '-qm', 'init')
        git(root, 'switch', '-qc', 'feat')

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.template, ignore_errors=True)

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix='harness-case-'))
        self.root = self.tmp / 'repo'
        shutil.copytree(self.template / 'repo', self.root, symlinks=True)
        self.bin = self.tmp / 'bin'
        self.bin.mkdir()
        self.env = {**os.environ, 'PATH': f'{self.bin}:{os.environ["PATH"]}', 'PYTHONDONTWRITEBYTECODE': '1',
                    'GOALRUN': GOALRUN}

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def h(self, *args, env=None):
        p = subprocess.run([sys.executable, 'scripts/harness', *args], cwd=self.root, capture_output=True,
                           text=True, env=env or self.env, timeout=600)
        return p.returncode, p.stdout + p.stderr

    def g(self, *args):
        p = subprocess.run([sys.executable, GOALRUN, *args], cwd=self.root, capture_output=True, text=True,
                           env=self.env, timeout=600)
        return p.returncode, p.stdout + p.stderr

    def write(self, name, text):
        (self.root / name).parent.mkdir(parents=True, exist_ok=True)
        (self.root / name).write_text(text)

    def commit(self, msg='work'):
        git(self.root, 'add', '-A')
        git(self.root, 'commit', '-q', '--allow-empty', '-m', msg)

    def body(self, text):
        path = self.tmp / 'body.md'
        path.write_text(text)
        return str(path)

    def proven_run(self, rid='F-001-r1', row=MUL_ROW, reqs='REQ-CALC-001\n', mul='a * b', test=TEST_MUL,
                   test_name='tests/test_mul.py'):
        """goalrun run with one row, code written, committed, whole-ledger verified on the commit."""
        self.assertEqual(self.g('init', rid, '--goal', 'g', '--spec', 'docs/spec.md')[0], 0)
        token = json.loads(self.g('resume', rid, '--owner', 't')[1])['token']
        gdir = self.root / '.testcases/runs' / rid / 'goalrun'
        (gdir / 'reqs.txt').write_text(reqs)
        (gdir / 'ledger.tsv').write_text(LEDGER_HEAD + row)
        if mul:
            self.write('src/calc.py', (self.root / 'src/calc.py').read_text().split('\n\ndef mul')[0]
                       + f'\n\ndef mul(a, b):\n    return {mul}\n')
        self.write(test_name, test)
        self.commit(rid)
        code, out = self.g('--run', rid, '--token', token, '--verify')
        self.assertEqual(code, 0, out)
        return git(self.root, 'rev-parse', 'HEAD')

    def fake_gh(self, head, body, login='owner', conclusion='success', earlier=None, pr_state=None):
        runs = [{'id': 2, 'name': 'verify', 'conclusion': conclusion, 'html_url': 'https://ci/2',
                 'app': {'slug': 'github-actions'}}]
        if earlier:
            runs.append({'id': 1, 'name': 'verify', 'conclusion': earlier, 'html_url': 'https://ci/1',
                         'app': {'slug': 'github-actions'}})
        routes = {
            'repos/{owner}/{repo}': {'full_name': 'o/r'},
            'repos/o/r/pulls/7': {'head': {'sha': head}, 'base': {'ref': 'main'}, **(pr_state or {'state': 'open'})},
            f'repos/o/r/commits/{head}/check-runs': {'check_runs': runs},
            # two pages: the approval sits past the first 100 comments
            'repos/o/r/issues/7/comments': [[{'user': {'login': 'bot'}, 'body': 'noise', 'html_url': 'x'}] * 100,
                                            [{'user': {'login': login}, 'created_at': 'x', 'body': body,
                                              'html_url': 'https://pr/7#c1'}]],
        }
        data = self.tmp / 'gh.json'
        data.write_text(json.dumps(routes))
        script = self.bin / 'gh'
        script.write_text('#!/usr/bin/env python3\nimport json,sys\n'
                          f'routes=json.load(open({str(data)!r}))\npath=sys.argv[-1]\n'
                          'for k,v in routes.items():\n'
                          '    if path.startswith(k):\n'
                          '        if isinstance(v,list) and "--slurp" not in sys.argv: v=v[0]\n'
                          '        print(json.dumps(v)); sys.exit(0)\n'
                          'sys.stderr.write("no route "+path); sys.exit(1)\n')
        script.chmod(0o755)

    def accept_f001(self, mul='a * b'):
        head = self.proven_run(mul=mul, row=MUL_ROW.replace('return a * b ::', f'return {mul} ::'))
        self.fake_gh(head, f'/accept F-001-r1 {head[:7]}')
        code, out = self.h('accept', 'F-001', '--pr', '7')
        self.assertEqual(code, 0, out)
        self.commit('receipt')


class Check(Fixture):
    def test_REQ_HAR_003_unknown_id_and_bad_filter_run_nothing(self):
        code, out = self.h('check', 'rm -rf /')
        self.assertEqual(code, 2, out)
        self.assertIn('unknown check id', out)
        code, out = self.h('check', 'unit', '--tests', 'x; touch pwned')
        self.assertEqual(code, 2, out)
        self.assertFalse((self.root / 'pwned').exists())
        self.assertIn('takes no --tests', self.h('check', 'docs', '--tests', 'x')[1])
        self.assertIn('pass --tests', self.h('check', 'one-file')[1])

    def test_REQ_HAR_004_missing_tool_and_no_tests_are_error(self):
        code, out = self.h('check', 'needs-missing')
        self.assertEqual(code, 2, out)
        self.assertIn('HARNESS-RESULT: ERROR needs-missing missing: no-such-tool-xyz', out)
        code, out = self.h('check', 'unit', '--tests', 'nothing_matches_this')
        self.assertEqual(code, 2, out)
        self.assertIn('no tests ran', out)

    def test_REQ_HAR_005_pass_fail_and_compile_error(self):
        self.assertEqual(self.h('check', 'unit')[0], 0)
        self.write('src/calc.py', 'def add(a, b):\n    return a - b - 0\n')  # new size: no stale .pyc
        self.assertEqual(self.h('check', 'unit')[0], 1)
        self.write('src/calc.py', 'def add(a, b):\n    return a +\n')
        self.assertEqual(self.h('check', 'unit')[0], 2)

    def test_REQ_HAR_005_go_test_json_classification(self):
        ev = lambda action, test=None, pkg='p': json.dumps({'Action': action, 'Package': pkg, 'Test': test})
        ok = '\n'.join([ev('run', 'TestA'), ev('pass', 'TestA'), ev('pass', 'TestB', 'q'), ev('pass')])
        self.assertEqual(H.classify('gotest', 0, ok, {}, 0), (H.PASS, '2 tests'))
        bad = '\n'.join([ev('pass', 'TestA'), ev('fail', 'TestB'), ev('fail', 'TestB/sub'), ev('fail')])
        self.assertEqual(H.classify('gotest', 1, bad, {}, 0)[0], H.FAIL)
        self.assertIn('TestB', H.classify('gotest', 1, bad, {}, 0)[1])
        # build error: plain text on stderr, a package-level fail, no failing test
        build = './x.go:3:1: syntax error\n' + ev('fail')
        self.assertEqual(H.classify('gotest', 1, build, {}, 0)[0], H.ERROR)
        # Go 1.10 vet or docker failure before any test ran
        self.assertEqual(H.classify('gotest', 2, 'vet: unreachable code', {}, 0)[0], H.ERROR)
        # a panic kills the binary: package fail without a test fail is a crash, not a caught defect
        self.assertEqual(H.classify('gotest', 1, ev('run', 'TestA') + '\npanic: boom\n' + ev('fail'), {}, 0)[0], H.ERROR)
        self.assertEqual(H.classify('gotest', 0, ev('skip', 'TestA') + '\n' + ev('pass'), {}, 0)[0], H.ERROR)
        self.assertEqual(H.classify('gotest', 0, 'not json {\n{broken', {}, 0)[0], H.ERROR)
        # core-go110: a build error under Go 1.10 is the finding itself
        self.assertEqual(H.classify('gotest', 1, build, {'build_error_fails': True}, 0)[0], H.FAIL)
        # but a toolchain container that never started (docker exit 125) is never a finding about the code
        self.assertEqual(H.classify('gotest', 125, 'Cannot connect to the Docker daemon', {'build_error_fails': True}, 0)[0], H.ERROR)
        self.assertEqual(H.classify('exit', 127, '', {}, 0)[0], H.ERROR)

    def test_REQ_HAR_005_real_go_test_pass_fail_and_build_error(self):
        prof = json.loads((self.root / '.harness/profile.json').read_text())
        prof['checks']['gounit'] = {'argv': ['go', 'test', '-json', './...'], 'kind': 'gotest', 'needs': ['go'],
                                    'paths': ['gox/*'], 'tests_template': ['-run', '{tests}'],
                                    'tests_filter': '^[A-Za-z0-9_]+$', 'cwd': 'gox', 'timeout': 300}
        self.write('.harness/profile.json', json.dumps(prof))
        self.write('gox/go.mod', 'module gox\n\ngo 1.20\n')
        self.write('gox/add.go', 'package gox\n\nfunc Add(a, b int) int { return a + b }\n')
        self.write('gox/add_test.go', 'package gox\n\nimport "testing"\n\n'
                   'func TestAdd(t *testing.T) {\n\tif Add(2, 3) != 5 {\n\t\tt.Fatal("add")\n\t}\n}\n')
        code, out = self.h('check', 'gounit')
        self.assertEqual(code, 0, out)
        self.assertIn('HARNESS-RESULT: PASS gounit 1 tests', out)
        self.assertEqual(self.h('check', 'gounit', '--tests', 'TestNothing')[0], 2)
        self.write('gox/add.go', 'package gox\n\nfunc Add(a, b int) int { return a - b }\n')
        self.assertEqual(self.h('check', 'gounit')[0], 1)
        self.write('gox/add.go', 'package gox\n\nfunc Add(a, b int) int { return a +\n')
        code, out = self.h('check', 'gounit')
        self.assertEqual(code, 2, out)

    def test_REQ_HAR_006_secrets_never_reach_or_leave_a_check(self):
        secret = 'abcd1234secretvalue'
        code, out = self.h('check', 'echo-secret', env={**self.env, 'FIXTURE_TOKEN': secret})
        self.assertEqual(code, 0, out)
        self.assertIn('env= pat=[redacted]', out)
        self.assertNotIn(secret, out)

    def test_REQ_HAR_034_ledger_cells_must_be_registry_calls(self):
        p = json.loads((self.root / '.harness/profile.json').read_text())
        self.assertIsNone(H.ledger_check_error('scripts/harness check unit --tests test_mul', p))
        self.assertIsNone(H.ledger_check_error('MANUAL:admin', p))
        for cell in ('scripts/harness check unit --tests test_mul; touch X', 'scripts/harness check unit && true',
                     'scripts/harness check unit --tests $(id)', 'python3 tests/run.py'):
            self.assertTrue(H.ledger_check_error(cell, p), cell)
        self.assertIn('kind exit', H.ledger_check_error('scripts/harness check docs', p))

    def test_REQ_HAR_040_crash_is_error(self):
        prof = json.loads((self.root / '.harness/profile.json').read_text())
        prof['checks']['unit'] = 'not a dict'
        self.write('.harness/profile.json', json.dumps(prof))
        self.assertEqual(self.h('check', 'unit')[0], 2)


class Corpus(Fixture):
    def test_REQ_HAR_030_031_rules_for_and_lint(self):
        self.write('.claude/rules/00-overview.md', '---\ndescription: map\n---\n# map\n')
        self.write('.claude/rules/01-calc.md', '---\ndescription: calc\npaths:\n  - "src/**"\n---\n# calc\n')
        self.write('.claude/skills/do-it/SKILL.md', '---\nname: do-it\ndescription: Use when x\n---\nUse skill `do-it`.\n')
        self.write('.claude/agents/helper.md', '---\nname: helper\ndescription: h\n---\nRun skill `do-it`.\n')
        self.commit('corpus')
        code, out = self.h('rules-for', 'src/calc.py')
        self.assertEqual(code, 0, out)
        self.assertIn('01-calc.md', out)
        self.assertIn('00-overview.md   (always)', out)
        self.assertEqual(self.h('rules-for', 'nope.py')[0], 2)
        self.assertEqual(self.h('lint-rules')[0], 0)
        self.write('.claude/rules/02-gone.md', '---\ndescription: g\npaths:\n  - "gone/**"\n---\nSee skill `missing`.\n')
        self.commit('stale rule')
        code, out = self.h('lint-rules')
        self.assertEqual(code, 1, out)
        self.assertIn('matches no tracked file', out)
        self.assertIn('unknown skill missing', out)


class SpecGate(Fixture):
    def test_REQ_HAR_035_product_change_needs_a_declared_feature_or_exemption(self):
        self.write('src/calc.py', CALC + '\n# tweak\n')
        self.commit('tweak')
        code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('no lines'))
        self.assertEqual(code, 1, out)
        self.assertIn('outside every `Feature:` line', out)
        code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file',
                           self.body('Spec-Exempt: comment only, no behaviour change\n'))
        self.assertEqual(code, 0, out)
        self.assertIn('SPEC-EXEMPT', out)
        code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('Feature: F-001\n'))
        self.assertEqual(code, 0, out)

    def test_REQ_HAR_035_new_revision_needs_a_spec_change(self):
        feats = json.loads((self.root / '.harness/features.json').read_text())
        feats['features'][0]['revision'] = 2
        self.write('.harness/features.json', json.dumps(feats))
        self.commit('bump')
        code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('Feature: F-001\n'))
        self.assertEqual(code, 1, out)
        self.assertIn('spec docs/spec.md is unchanged', out)

    def test_REQ_HAR_036_accepted_spec_is_frozen(self):
        self.accept_f001()
        git(self.root, 'branch', '-f', 'main', 'HEAD')
        self.write('docs/spec.md', (self.root / 'docs/spec.md').read_text() + 'REQ-CALC-002 zero.\n')
        self.commit('edit spec')
        code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('Feature: F-001\n'))
        self.assertEqual(code, 1, out)
        self.assertIn('changed after r1 was accepted', out)
        feats = json.loads((self.root / '.harness/features.json').read_text())
        feats['features'][0]['revision'] = 2
        self.write('.harness/features.json', json.dumps(feats))
        self.commit('r2')
        self.assertEqual(self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('Feature: F-001\n'))[0], 0)

    def test_REQ_HAR_036_037_moving_the_spec_path_needs_a_revision(self):
        self.accept_f001()
        self.write('docs/other.md', '# Other\n')  # already on main: the PR only repoints `spec`
        self.commit('other doc')
        git(self.root, 'branch', '-f', 'main', 'HEAD')
        feats = json.loads((self.root / '.harness/features.json').read_text())
        feats['features'][0]['spec'] = 'docs/other.md'
        self.write('.harness/features.json', json.dumps(feats))
        self.commit('move spec')
        for line in ('Feature: F-001\n', 'Feature: F-001 bug x\n'):
            code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body(line))
            self.assertEqual(code, 1, (line, out))
        self.assertIn('spec moved docs/spec.md → docs/other.md', out)

    def test_REQ_HAR_036_deleting_a_receipt_does_not_thaw_the_spec(self):
        self.accept_f001()
        git(self.root, 'branch', '-f', 'main', 'HEAD')
        shutil.rmtree(self.root / '.harness/acceptance')
        self.write('docs/spec.md', (self.root / 'docs/spec.md').read_text() + 'REQ-CALC-002 zero.\n')
        self.commit('drop receipt, edit spec')
        code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('Feature: F-001\n'))
        self.assertEqual(code, 1, out)
        self.assertIn('append-only', out)

    def test_REQ_HAR_037_bug_line_keeps_spec_and_needs_acceptance(self):
        self.write('src/calc.py', CALC + '\n# fix\n')
        self.commit('fix')
        code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('Feature: F-001 bug neg\n'))
        self.assertEqual(code, 1, out)
        self.assertIn('no committed receipt', out)

    def test_REQ_HAR_038_forged_receipt_fails(self):
        r = {k: 'x' for k in H.RECEIPT_FIELDS}
        r.update(feature='F-001', revision=1, run_id='F-001-r1', approval={'login': 'someone-else'},
                 head_sha=git(self.root, 'rev-parse', 'HEAD'))
        self.write('.harness/acceptance/F-001/F-001-r1-x.json', json.dumps(r))
        self.commit('forged')
        code, out = self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('Spec-Exempt: x\n'))
        self.assertEqual(code, 1, out)
        self.assertIn('not from an approver', out)
        self.write('.harness/acceptance/F-001/F-001-r1-x.json', '{broken')
        self.commit('broken')
        self.assertIn('not valid JSON', self.h('spec-gate', '--base', 'main', '--pr-body-file', self.body('Spec-Exempt: x\n'))[1])


class Accept(Fixture):
    def test_REQ_HAR_039_merged_pr_still_accepts_but_closed_unmerged_does_not(self):
        head = self.proven_run()
        self.fake_gh(head, f'/accept F-001-r1 {head[:7]}', pr_state={'state': 'closed', 'merged_at': None})
        code, out = self.h('accept', 'F-001', '--pr', '7')
        self.assertEqual(code, 2, out)
        self.assertIn('open or merged', out)
        # merged, but neither the head nor a merge commit is in local main yet (not fetched)
        git(self.root, 'update-ref', 'refs/remotes/origin/main', 'HEAD~1')
        self.fake_gh(head, f'/accept F-001-r1 {head[:7]}', pr_state={'state': 'closed', 'merged_at': '2026-11-03T05:46:47Z'})
        code, out = self.h('accept', 'F-001', '--pr', '7')
        self.assertEqual(code, 2, out)
        self.assertIn('git fetch origin first', out)
        git(self.root, 'update-ref', 'refs/remotes/origin/main', head)
        base = git(self.root, 'rev-parse', 'HEAD~1')
        self.fake_gh(head, f'/accept F-001-r1 {head[:7]}',
                     pr_state={'state': 'closed', 'merged_at': '2026-11-03T05:46:47Z', 'base': {'ref': 'main', 'sha': base}})
        code, out = self.h('accept', 'F-001', '--pr', '7')
        self.assertEqual(code, 0, out)
        self.assertIn('new PR', out)
        # no merge commit known locally: the receipt keeps the PR's base commit, never the head itself
        receipt = json.loads(next((self.root / '.harness/acceptance/F-001').glob('*.json')).read_text())
        self.assertEqual(receipt['base_sha'], base)
        self.assertNotEqual(receipt['base_sha'], receipt['head_sha'])

    def test_REQ_HAR_017_039_owner_merge_is_the_approval_and_squash_counts(self):
        head = self.proven_run()
        branch = git(self.root, 'rev-parse', '--abbrev-ref', 'HEAD')
        # a squash merge: one new commit on main beside the PR head, which itself never reaches main
        git(self.root, 'checkout', '-q', '-b', 'squash', 'HEAD~1')
        git(self.root, 'checkout', '-q', head, '--', '.')
        git(self.root, 'commit', '-q', '-m', 'squash of PR 7')
        squash = git(self.root, 'rev-parse', 'HEAD')
        git(self.root, 'checkout', '-q', branch)
        git(self.root, 'update-ref', 'refs/remotes/origin/main', squash)
        merged = {'state': 'closed', 'merged_at': '2026-11-05T10:00:00Z', 'merge_commit_sha': squash,
                  'html_url': 'https://pr/7'}
        # merged by someone who is not an approver, no /accept comment: no receipt
        self.fake_gh(head, 'looks good', pr_state={**merged, 'merged_by': {'login': 'someone-else'}})
        code, out = self.h('accept', 'F-001', '--pr', '7')
        self.assertEqual(code, 2, out)
        self.assertFalse((self.root / '.harness/acceptance').exists())
        # merged by the owner: the merge is the approval, no comment needed
        self.fake_gh(head, 'looks good', pr_state={**merged, 'merged_by': {'login': 'owner'}})
        code, out = self.h('accept', 'F-001', '--pr', '7')
        self.assertEqual(code, 0, out)
        receipt = json.loads(next((self.root / '.harness/acceptance/F-001').glob('*.json')).read_text())
        self.assertEqual((receipt['approval']['login'], receipt['approval']['merge_commit']), ('owner', squash))
        self.assertEqual(receipt['head_sha'], head)
        # on main the squash commit stands in for the head, so the receipt counts there
        receipt['_path'] = self.root / '.harness/acceptance/F-001/x.json'
        old = H.ROOT
        H.ROOT = self.root
        try:
            git(self.root, 'stash', '-q', '-u')
            git(self.root, 'checkout', '-q', 'squash')
            self.assertNotIn('not an ancestor', ' '.join(H.receipt_errors(receipt, {'approvers': ['owner']},
                                                                            {'F-001': {}})))
            receipt['approval']['merge_commit'] = None
            self.assertIn('not an ancestor', ' '.join(H.receipt_errors(receipt, {'approvers': ['owner']},
                                                                         {'F-001': {}})))
        finally:
            H.ROOT = old

    def test_REQ_HAR_017_039_no_approval_no_receipt(self):
        head = self.proven_run()
        sha = head[:7]
        for body, login, ci, earlier in (('looks good', 'owner', 'success', None),
                                         (f'/accept F-001-r1 {sha}', 'someone-else', 'success', None),
                                         (f'/accept F-001-r1 {sha}', 'owner', 'failure', 'success'),
                                         ('/accept F-001-r1 0000000', 'owner', 'success', None),
                                         ('/accept F-001-r1', 'owner', 'success', None)):
            self.fake_gh(head, body, login, ci, earlier)
            code, out = self.h('accept', 'F-001', '--pr', '7')
            self.assertEqual(code, 2, (body, login, ci, out))
            self.assertFalse((self.root / '.harness/acceptance').exists())

    def test_REQ_HAR_018_receipt_fields(self):
        head = self.proven_run()
        self.fake_gh(head, f'/accept F-001-r1 {head[:7]}')
        code, out = self.h('accept', 'F-001', '--pr', '7')
        self.assertEqual(code, 0, out)
        receipt = json.loads(next((self.root / '.harness/acceptance/F-001').glob('*.json')).read_text())
        for key in H.RECEIPT_FIELDS:
            self.assertIn(key, receipt)
        self.assertEqual((receipt['head_sha'], receipt['reqs']), (head, ['REQ-CALC-001']))

    def test_REQ_HAR_012_stale_proof_is_refused(self):
        self.proven_run()
        self.write('src/calc.py', (self.root / 'src/calc.py').read_text() + '# later\n')
        self.commit('later')
        self.fake_gh(git(self.root, 'rev-parse', 'HEAD'), '/accept F-001-r1 x')
        self.assertEqual(self.h('accept', 'F-001', '--pr', '7')[0], 3)

    def test_REQ_HAR_013_break_that_crashes_is_error(self):
        row = MUL_ROW.replace('return a * b :: return a + b', 'return a * b :: return a *')
        head = self.proven_run(row=row)
        self.fake_gh(head, f'/accept F-001-r1 {head[:7]}')
        code, out = self.h('accept', 'F-001', '--pr', '7')
        self.assertEqual(code, 2, out)
        self.assertIn('crashed the check', out)

    def test_REQ_HAR_037_bug_run_cannot_add_a_requirement(self):
        self.accept_f001(mul='abs(a) * b')  # mul(-2, 3) gives 6
        neg = (TEST_MUL.replace('Mul', 'Neg').replace('test_mul', 'test_mul_neg')
               .replace('calc.mul(2, 3), 6', 'calc.mul(-2, 3), -6'))
        row = ('NEG\tREQ-CALC-001 mul keeps the sign\tscripts/harness check unit --tests test_mul_neg\tsrc/calc.py\t'
               'src/calc.py :: return a * b :: return abs(a) * b\n')
        head = self.proven_run('F-001-r1-bug-neg', row=row, mul='a * b', test=neg, test_name='tests/test_neg.py')
        self.fake_gh(head, f'/accept F-001-r1-bug-neg {head[:7]}')
        code, out = self.h('accept', 'F-001', '--bug', 'neg', '--pr', '7')
        self.assertEqual(code, 0, out)
        self.commit('bug receipt')
        zero = ('ZERO\tREQ-CALC-009 zero rule\tscripts/harness check unit --tests test_mul\t—\t'
                'src/calc.py :: return a * b :: return a + b\n')
        head = self.proven_run('F-001-r1-bug-zero', row=zero, reqs='REQ-CALC-009\n', mul=None, test=TEST_MUL)
        self.fake_gh(head, f'/accept F-001-r1-bug-zero {head[:7]}')
        code, out = self.h('accept', 'F-001', '--bug', 'zero', '--pr', '7')
        self.assertEqual(code, 1, out)
        self.assertIn('spec was silent', out)


class Guard(unittest.TestCase):
    def decide(self, tool, **inp):
        return G.decide({'tool_name': tool, 'tool_input': inp})

    def test_REQ_HAR_033_denies_master_push_and_merge(self):
        for cmd in ('git push origin master', 'git push origin main', 'rtk git push origin HEAD:master',
                    'gh pr merge 5', 'git -C x push origin master', 'cd x && git push origin +refs/heads/master',
                    'gh api -X PUT repos/o/r/pulls/5/merge', 'git push origin \\\nmaster',
                    'git push -u origin \\\n  HEAD:refs/heads/master'):
            self.assertEqual(self.decide('Bash', command=cmd)[0], 'deny', cmd)

    def test_REQ_HAR_033_bare_push_on_main_is_denied(self):
        repo = Path(tempfile.mkdtemp(prefix='guard-'))
        self.addCleanup(shutil.rmtree, repo, True)
        init_repo(repo)
        for cmd in ('git push', 'git push origin HEAD'):
            self.assertEqual(G.decide({'tool_name': 'Bash', 'tool_input': {'command': cmd}, 'cwd': str(repo)})[0], 'deny', cmd)
        git(repo, 'switch', '-qc', 'feat/x')
        for cmd in ('git push', 'git push origin HEAD', f'git -C {repo}/missing push'):
            self.assertIsNone(G.decide({'tool_name': 'Bash', 'tool_input': {'command': cmd}, 'cwd': str(repo)}), cmd)

    def test_REQ_HAR_033_merge_is_denied_only_on_main(self):
        repo = Path(tempfile.mkdtemp(prefix='guard-'))
        self.addCleanup(shutil.rmtree, repo, True)
        init_repo(repo)
        on = lambda cmd: G.decide({'tool_name': 'Bash', 'tool_input': {'command': cmd}, 'cwd': str(repo)})
        for cmd in ('git merge feat/x', 'git merge --no-edit origin/main', f'git -C {repo} merge feat/x'):
            self.assertEqual(on(cmd)[0], 'deny', cmd)
        git(repo, 'switch', '-qc', 'feat/x')
        for cmd in ('git merge --no-edit origin/main', 'git merge main'):
            self.assertIsNone(on(cmd), cmd)

    def test_REQ_HAR_033_asks_before_public_actions(self):
        for cmd in ('gh release create v2.0.0', 'gh release upload v2.0.0 x.exe', 'git push origin v2.0.0',
                    'git push --tags', 'git push origin refs/tags/v2.0.0', 'git push -f origin feat/x',
                    'git push --force-with-lease origin feat/x', 'gh issue comment 7 --body hi', 'gh issue close 1',
                    'gh pr close 5', 'gh pr comment 5 --body x', 'gh api -X POST repos/o/r/issues/1/comments',
                    'gh api repos/o/r/issues/1/comments -f body=x', 'gh repo edit --visibility private'):
            self.assertEqual(self.decide('Bash', command=cmd)[0], 'ask', cmd)

    def test_REQ_HAR_033_allows_ordinary_work(self):
        for cmd in ('git push -u origin feat/harness', 'scripts/harness check go-test', 'go test ./...',
                    "cat > rule.md <<'EOF'\ngit push origin master\ngh release create v9\nEOF",
                    'git push -u origin feat/master-menu', 'python3 -c "print(\'git push origin master\')"',
                    'git log origin/master..HEAD', 'gh pr view 5', 'gh pr create --base master --body-file b.md',
                    'gh issue list', 'gh api repos/o/r/issues/1/comments', 'gh release list'):
            self.assertIsNone(self.decide('Bash', command=cmd), cmd)
        self.assertIsNone(self.decide('Read', file_path='/x/server.go'))

    def test_REQ_HAR_033_git_pre_push_refuses_main_for_every_client(self):
        hook = REPO / '.githooks/pre-push'
        p = subprocess.run([str(hook), 'origin', 'url'], input='refs/heads/x a refs/heads/master b\n',
                           capture_output=True, text=True)
        self.assertEqual(p.returncode, 1, p.stderr)
        self.assertIn('forbidden', p.stderr)
        self.assertTrue((REPO / '.agents/skills').resolve() == (REPO / '.claude/skills').resolve())


class Workflow(unittest.TestCase):
    def test_REQ_HAR_034_glob_filter_is_allowed_only_when_the_shell_leaves_it_alone(self):
        # goalrun runs the cell with `sh -c`: an unquoted `*` stays literal only if it matches no file there
        prof = json.loads((REPO / '.harness/profile.json').read_text())
        self.assertIsNone(H.ledger_check_error('scripts/harness check unit --tests *Mul*', {'checks': {'unit': {'kind': 'unittest', 'tests_filter': '^[A-Za-z0-9_.*]+$'}}}))
        self.assertIn('matches files', H.ledger_check_error('scripts/harness check unit --tests *.md', {'checks': {'unit': {'kind': 'unittest', 'tests_filter': '^[A-Za-z0-9_.*]+$'}}}))
        self.assertIsNone(H.ledger_check_error('scripts/harness check go-test --tests TestStringNamesVersionAndEdition', prof))
        self.assertIn('kind exit', H.ledger_check_error('scripts/harness check go-vet', prof))
        self.assertIn('must be', H.ledger_check_error('foo bar baz --tests *.md', prof))

    def test_REQ_HAR_018_merged_receipt_base_is_the_fork_point(self):
        repo = Path(tempfile.mkdtemp(prefix='fork-'))
        self.addCleanup(shutil.rmtree, repo, True)
        init_repo(repo)
        for name in ('a', 'b'):
            (repo / name).write_text(name)
            git(repo, 'add', name)
            git(repo, 'commit', '-qm', name)
        fork = git(repo, 'rev-parse', 'HEAD')
        git(repo, 'switch', '-qc', 'feat')
        (repo / 'f').write_text('f')
        git(repo, 'add', 'f')
        git(repo, 'commit', '-qm', 'f')
        head = git(repo, 'rev-parse', 'HEAD')
        git(repo, 'switch', '-q', 'main')
        (repo / 'c').write_text('c')
        git(repo, 'add', 'c')
        git(repo, 'commit', '-qm', 'c')
        git(repo, 'merge', '-q', '--no-ff', '-m', 'merge', 'feat')
        mc = git(repo, 'rev-parse', 'HEAD')
        run = lambda *args: subprocess.run(['git', *args], cwd=repo, capture_output=True, text=True).stdout.strip()
        # the merge commit is known: fork point, not main's tip at PR time, not the head
        self.assertEqual(H.merged_base_sha({'merge_commit_sha': mc, 'base': {'sha': 'f' * 40}}, head, run), fork)
        # merge commit not known locally: the PR's own base commit
        self.assertEqual(H.merged_base_sha({'merge_commit_sha': '0' * 40, 'base': {'sha': 'e' * 40}}, head, run), 'e' * 40)
        self.assertIsNone(H.merged_base_sha({}, head, run))

    def test_REQ_HAR_022_verify_workflow_runs_every_gate(self):
        text = (REPO / '.github/workflows/verify.yml').read_text()
        for needle in ('pull_request', 'scripts/harness lint-rules', 'scripts/harness spec-gate', 'PR_BODY',
                       'scripts/harness check', 'GOALRUN', '\n  verify:', 'setup-go'):
            self.assertIn(needle, text)
        self.assertFalse((REPO / '.harness/vendor').exists(), 'goalrun must not be copied into the repo')

    def test_REQ_HAR_022_affected_checks_match_the_harness(self):
        import textwrap
        text = (REPO / '.github/workflows/verify.yml').read_text()
        snippet = textwrap.dedent(text.split("python3 -c '", 1)[1].split("')", 1)[0])
        prof = json.loads((REPO / '.harness/profile.json').read_text())
        for files in (['internal/core/buildinfo/buildinfo.go'], ['cmd/printerone/main.go'], ['README.md'],
                      ['docs/x.md'], ['scripts/go110', 'go.mod'], ['.claude/rules/00-project-overview.md']):
            out = subprocess.run([sys.executable, '-c', snippet], cwd=REPO, input='\n'.join(files),
                                 capture_output=True, text=True, check=True).stdout
            want = {c for c, d in prof['checks'].items() if not d.get('tests_required') and
                    any(H.matches(f, d.get('paths', [])) for f in files)} | set(prof.get('always_checks', []))
            self.assertEqual(set(json.loads(out)), want, files)


if __name__ == '__main__':
    unittest.main()
