# Agent Guidelines

## Core Principles

- Do not maintain backward compatibility unless explicitly requested.
- Follow the existing document and script patterns before adding new structure.
- Keep generated PPTX artifacts out of git unless the user explicitly asks.

## Project Overview

- Project type: PPTX design review guideline and tooling repository.
- Primary languages: YAML, Markdown, Python, JavaScript tooling.
- Main guideline: `doc/slide-guideline-v1.yml`.
- Task ledger: `doc/tasks.md`.

## Commands

```bash
# Validate guideline YAML
node -e "const fs=require('fs'); const yaml=require('js-yaml'); \
yaml.load(fs.readFileSync('doc/slide-guideline-v1.yml','utf8')); \
console.log('YAML OK')"

# Markdown lint
npx --no-install markdownlint-cli2 doc/tasks.md AGENTS.md TASKS.md

# PPTX lint regression
python3 skills/pptx-design-reviewer/scripts/test_pptx_lint.py
```

## Commit Rules

- Use Japanese Conventional Commit subject and body.
- Use `docs(rules)` for guideline and task-ledger rule changes.
- Run the relevant commands above before committing.

## Architecture

- `doc/`: design-system guideline, task ledger, and review notes.
- `skills/pptx-design-reviewer/`: PPTX lint, fix, repair, and review scripts.
- `skills/web-design-reviewer/`: web design review skill references.
