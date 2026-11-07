#!/usr/bin/env python3
"""PreToolUse guard for agent sessions in this repo (registered in .claude/settings.json).

Denies what no agent may do here (push or merge master/main); asks the owner before anything public: a release,
a tag push, a force push, a comment or state change on an issue or someone else's PR. Everything else passes.
Reads the hook JSON on stdin, prints a Claude Code permission decision, exits 0.

Bash commands are split into simple commands (shell quoting respected, heredoc bodies dropped), so a branch
named feat/master-x or a string that merely mentions `git push … master` is not a push to master.
A guard, not a sandbox: the host's permission system still applies. Known gaps: text run through `bash -c`,
`eval` or `find -exec` is data to this parser and passes; `.githooks/pre-push` still refuses master/main.
Adapted from the LaunchPadTV guard; spec docs/specs/2026-11-07-harness.md (REQ-HAR-033).
"""
import json
import os
import re
import shlex
import subprocess
import sys

PREFIXES = {'rtk', 'command', 'env', 'sudo', 'nohup', 'time', 'exec'}
MAIN = {'main', 'master'}
SEPARATORS = {';', '&&', '||', '|', '&', '(', ')', '|&', ';;'}
# gh subcommands that publish or change shared state on GitHub; `gh pr create` on a branch is ordinary work
GH_ASK = {
    ('release', 'create'): 'publishes a GitHub release', ('release', 'upload'): 'changes a GitHub release',
    ('release', 'edit'): 'changes a GitHub release', ('release', 'delete'): 'deletes a GitHub release',
    ('issue', 'comment'): 'posts on a public issue', ('issue', 'close'): 'closes a public issue',
    ('issue', 'edit'): 'edits a public issue', ('issue', 'create'): 'opens a public issue',
    ('issue', 'reopen'): 'reopens a public issue', ('issue', 'lock'): 'locks a public issue',
    ('pr', 'comment'): 'posts on a public PR', ('pr', 'close'): 'closes a PR', ('pr', 'review'): 'reviews a PR',
    ('repo', 'edit'): 'changes repository settings', ('repo', 'delete'): 'deletes a repository',
}


def drop_heredocs(cmd):
    """Remove heredoc bodies: their text is data for the command, never a command itself."""
    out, lines, i = [], cmd.split('\n'), 0
    while i < len(lines):
        out.append(lines[i])
        tags = re.findall(r'<<-?\s*[\'"]?([A-Za-z_][A-Za-z0-9_]*)[\'"]?', lines[i])
        i += 1
        for tag in tags:
            while i < len(lines) and lines[i].strip() != tag:
                i += 1
            i += 1
    return '\n'.join(out)


def simple_commands(cmd):
    """argv lists of the simple commands in a shell line; None if it cannot be parsed."""
    lexer = shlex.shlex(drop_heredocs(cmd).replace('\\\n', ' ').replace('\n', ' ; '), posix=True, punctuation_chars=True)
    lexer.whitespace_split = True
    try:
        tokens = list(lexer)
    except ValueError:
        return None
    cmds, cur = [], []
    for tok in tokens:
        if tok in SEPARATORS:
            cmds.append(cur)
            cur = []
        else:
            cur.append(tok)
    cmds.append(cur)
    out = []
    for argv in cmds:
        while argv and (argv[0] in PREFIXES or re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*=.*', argv[0])):
            argv = argv[1:]
        if argv:
            out.append(argv)
    return out


def git_sub(argv):
    """(subcommand, args after it, -C dir) for a git argv."""
    i, cwd = 1, None
    while i < len(argv) and argv[i].startswith('-'):
        if argv[i] in ('-C', '-c', '--git-dir', '--work-tree'):
            cwd = argv[i + 1] if argv[i] == '-C' and i + 1 < len(argv) else cwd
            i += 2
        else:
            i += 1
    return (argv[i], argv[i + 1:], cwd) if i < len(argv) else (None, [], cwd)


def current_branch(cwd):
    try:
        p = subprocess.run(['git', 'symbolic-ref', '--short', '-q', 'HEAD'], cwd=cwd or None, capture_output=True, text=True)
    except OSError:  # -C to a missing dir: git itself fails the push; the hook must still exit 0
        return ''
    return p.stdout.strip() if p.returncode == 0 else ''


def names_main(ref):
    dst = ref.lstrip('+').split(':')[-1]
    return re.sub(r'^(refs/heads/|origin/)', '', dst) in MAIN


def names_tag(ref):
    return ref.lstrip('+').split(':')[-1].startswith('refs/tags/')


def check_argv(argv, cwd):
    name = os.path.basename(argv[0])
    if name == 'git':
        sub, args, gcwd = git_sub(argv)
        here = os.path.join(cwd or '.', gcwd) if gcwd else cwd
        if sub == 'push':
            pos = [a for a in args if not a.startswith('-')]
            on_main = len(pos) <= 1 or any(r.lstrip('+').split(':')[-1] == 'HEAD' for r in pos[1:])
            if any(names_main(r) for r in pos[1:]) or (on_main and current_branch(here) in MAIN):
                return 'deny', 'push to master/main is forbidden: push a branch and open a PR (CLAUDE.md)'
            if any(a in ('-f', '--force') or a.startswith('--force-with-lease') or
                   (a.startswith('-') and not a.startswith('--') and 'f' in a) for a in args):
                return 'ask', 'force push'
            # a tag push triggers the release workflow
            if '--tags' in args or any(names_tag(r) or re.fullmatch(r'v\d[\w.-]*', r) for r in pos[1:]):
                return 'ask', 'pushes a tag (starts a release)'
        # merging INTO master is the owner's job; merging master into a feature branch is ordinary work
        if sub == 'merge' and current_branch(here) in MAIN:
            return 'deny', 'merging into master is the owner\'s job (CLAUDE.md)'
    if name == 'gh':
        if argv[1:3] == ['pr', 'merge']:
            return 'deny', 'merging is the owner\'s job (CLAUDE.md)'
        if len(argv) > 1 and argv[1] == 'api':
            if any(re.search(r'pulls/\d+/merge\b', a) for a in argv[2:]):
                return 'deny', 'merging is the owner\'s job (CLAUDE.md)'
            method = next((argv[i + 1] for i, a in enumerate(argv[:-1]) if a in ('-X', '--method')), None)
            method = method or next((a[2:] for a in argv if a.startswith('-X') and len(a) > 2), None)
            writes = (method or '').upper() not in ('', 'GET') or any(
                a in ('-f', '-F', '--field', '--raw-field', '--input') for a in argv)
            if writes:
                return 'ask', 'gh api write to GitHub'
        why = GH_ASK.get(tuple(argv[1:3]))
        if why:
            return 'ask', why
    return None


def decide(event):
    if event.get('tool_name') != 'Bash':
        return None
    cmds = simple_commands(str((event.get('tool_input') or {}).get('command', '')))
    if cmds is None:
        return 'ask', 'command could not be parsed for the repo guard'
    verdicts = [v for v in (check_argv(argv, event.get('cwd')) for argv in cmds) if v]
    return next((v for v in verdicts if v[0] == 'deny'), verdicts[0] if verdicts else None)


def main():
    try:
        event = json.load(sys.stdin)
    except json.JSONDecodeError:
        return 0  # never block on a malformed event; the permission system still applies
    verdict = decide(event)
    if verdict:
        decision, why = verdict
        print(json.dumps({'hookSpecificOutput': {
            'hookEventName': 'PreToolUse', 'permissionDecision': decision,
            'permissionDecisionReason': f'harness guard: {why}'}}))
    return 0


if __name__ == '__main__':
    sys.exit(main())
