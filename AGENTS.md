# Agent instructions (Codex and other agents)

`CLAUDE.md` is the one source of project instructions; read it first. This file only adds what a non-Claude agent
does differently.

- Rules are not loaded for you. Before reading or editing any file run
  `scripts/harness rules-for <path>` (or pipe `git diff --name-only origin/master...HEAD` into it) and read every
  file it prints.
- Skills live in `.claude/skills/<name>/SKILL.md`; `.agents/skills` links there. Use the one `CLAUDE.md` names for
  the task and follow it.
- Roles in `.claude/agents/` are briefs you can follow when you are given that role.
- Hooks in `.claude/settings.json` do not run for you: still never push or merge `master`, never release, tag or
  comment on GitHub without the owner's approval. Git hook `.githooks/pre-push` refuses a push to
  `master`/`main` once a clone runs `git config core.hooksPath .githooks`.
- Reviews follow skill `review-change`; findings in English with `file:line` and P0–P3.
