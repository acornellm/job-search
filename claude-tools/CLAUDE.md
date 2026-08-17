# Job Search Claude Skills

**Repository**: https://github.com/acornellm/job-search
**Owner**: Adam Myers | acornellm

Production workflow skills for Claude Code CLI. Each skill guides Claude through a recipe to produce tangible output — not knowledge dumps, but working deliverables.

## Philosophy

- Every skill must produce visible output (files, configurations, deployable projects)
- "The context window is a public good" — only include what Claude doesn't already know
- **Teach patterns, not ship scripts** — skills teach Claude *what* to do, Claude generates scripts adapted to the user's environment. Pre-built scripts in `scripts/` are the rare exception, not the default. Put proven implementation patterns in `references/` for Claude to adapt.
- Follows the official Claude Code plugin spec

## Directory Structure

```
claude-tools/
├── plugins/                                # one folder per plugin; the folders are the truth, don't keep counts here
│   └── writing/                            # professional documents
│       └── skills/
│           └── resume-cover-letter/
├── CLAUDE.md                               # This file
└── README.md                               # Public-facing overview
```

## Plugin Anatomy (Anthropic Spec)

Each plugin contains one or more skills, auto-discovered from `skills/`:

```
plugin-name/
├── .claude-plugin/
│   └── plugin.json        # name, description, author
└── skills/
    └── skill-name/
        ├── SKILL.md       # Frontmatter + instructions (inline everything critical)
        ├── ERRATA.md      # Optional: versioned corrections discovered during builds
        ├── scripts/       # Executable scripts the agent RUNS (not reads)
        ├── references/    # Supplementary/variant docs (NOT critical path)
        └── assets/        # Files used in output (templates, images)
```

### Skill Design: Inline Everything Critical

**If the agent skipping it would derail the workflow, it goes in SKILL.md.** Reference files are for genuinely optional material — variant-specific docs, supplementary examples, historical context. Anything on the critical path must be inline.

This was learned the hard way: an agent was told "see references/stitch-direct.md for the curl commands." It skipped the file entirely and tried to use the website in a browser instead. The critical commands were 20 lines away in a reference file. It never read them.

| Content type | Where it goes | Example |
|-------------|--------------|---------|
| Workflow steps, commands, scripts | **SKILL.md body (inline)** | curl commands, Python scripts, mapping tables |
| Variant/optional docs | `references/` | Platform-specific variants (AWS vs GCP) |
| Templates copied into user projects | `assets/` | React boilerplate, config files |

**Why not reference files for critical content?** When a skill loads, SKILL.md goes directly into context. The agent sees it immediately. Reference files require a deliberate choice to read another file — an extra decision point that LLMs deprioritise in favour of acting. The instruction to "go read file X" competes with the instruction to "do the task" and loses.

**No file size anxiety.** The old 500-line limit was a context economics rule from the 200K era. A 500-line skill is ~2500 tokens — 0.25% of 1M context, 1.25% of 200K. Even on smaller contexts, a working skill that's 800 lines beats a broken skill that's 300 lines with critical content in references the agent never reads.

### Frontmatter Validation

- `name`: kebab-case, lowercase letters/digits/hyphens, max 64 characters
- `description`: max 1024 characters, no angle brackets. Include trigger phrases.
- Optional: `license`, `compatibility`, `allowed-tools`, `metadata`

## Quality Bar

Before committing a skill:
- [ ] SKILL.md has valid YAML frontmatter (name: kebab-case max 64 chars, description: max 1024 chars)
- [ ] Everything on the critical path is inline in SKILL.md (no "see references/" for must-do steps)
- [ ] Produces tangible output (not just reference material)
- [ ] Tested by actually using it on a real task
- [ ] Rich enough that the agent doesn't need to improvise — include exact commands, scripts, mapping tables
- [ ] Not brutally summarised — detail is better than brevity when the detail prevents mistakes

## Skill Errata (ERRATA.md)

When a skill's instructions are correct at one point but a library update changes behaviour, capture the correction in `ERRATA.md` alongside the SKILL.md rather than immediately rewriting the skill.

**Status lifecycle**: `active` (current correction) → `absorbed` (folded into SKILL.md) → `outdated` (library changed again)

Only for version-specific issues. Small typos or obvious mistakes should just be fixed in SKILL.md directly.


