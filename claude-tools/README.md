# claude-tools

Job-search Skills for Claude. See `CLAUDE.md` for the authoring spec — this
file covers how to *install and use* a finished skill once it's built.

Currently shipped:

- `plugins/writing/skills/resume-cover-letter` — tailored resume and cover
  letter writer (ATS formatting, CAR bullets, tone matching, etc.)

## Package a skill for install

`SKILL.md` (plus any `scripts/`, `references/`, `assets/`) needs to be zipped
with `SKILL.md` at the archive root — not nested under the skill's folder
name — for either install path below.

```bash
claude-tools/scripts/package-skill.sh claude-tools/plugins/writing/skills/resume-cover-letter
# -> claude-tools/dist/resume-cover-letter.zip
```

`dist/` is build output — add it to `.gitignore` if you don't want the zip
committed.

## Install to the Claude desktop app

1. Open **Settings → Capabilities**.
2. Turn on **Code execution** and **File creation** — skills run inside the
   code-execution sandbox, so both are required or the skill won't load.
3. Under **Skills**, click **Add**, then **Upload a skill**, and pick either
   the zip from `dist/` or the skill folder directly (the app accepts both).
4. Toggle the skill on. Claude reads the `description` from the frontmatter
   to decide when to invoke it — mention "resume", "cover letter", or a job
   application to trigger `resume-cover-letter`.

Same steps work on claude.ai in a browser (Settings → Capabilities → Skills).

## Install to the Claude API

Skills on the API are a Messages API feature — Claude runs them via the
code-execution tool, so every request that uses one needs the code-execution
tool declared and the `skills-2025-10-02` + `code-execution-2025-08-25` beta
headers. This is a different call shape than a plain `messages.create`.

### 1. Upload the skill (one-time, or per update)

```bash
curl -X POST "https://api.anthropic.com/v1/skills" \
  -H "x-api-key: $ANTHROPIC_API_KEY" \
  -H "anthropic-version: 2023-06-01" \
  -H "anthropic-beta: skills-2025-10-02" \
  -F "files[]=@claude-tools/plugins/writing/skills/resume-cover-letter/SKILL.md;filename=resume-cover-letter/SKILL.md"
```

Response returns a generated `id` (e.g. `skill_01AbCdEfGhIjKlMnOpQrStUv`) —
save it, you'll reference it by that ID, not by name.

```json
{
  "id": "skill_01AbCdEfGhIjKlMnOpQrStUv",
  "display_title": null,
  "latest_version": "1759178010641129",
  "source": "custom",
  "type": "skill"
}
```

### 2. Push a new version after editing SKILL.md

```bash
curl -X POST "https://api.anthropic.com/v1/skills/skill_01AbCdEfGhIjKlMnOpQrStUv/versions" \
  -H "x-api-key: $ANTHROPIC_API_KEY" \
  -H "anthropic-version: 2023-06-01" \
  -H "anthropic-beta: skills-2025-10-02" \
  -F "files[]=@claude-tools/plugins/writing/skills/resume-cover-letter/SKILL.md;filename=resume-cover-letter/SKILL.md"
```

### 3. Use it in a Messages API call

```bash
curl https://api.anthropic.com/v1/messages \
  -H "x-api-key: $ANTHROPIC_API_KEY" \
  -H "anthropic-version: 2023-06-01" \
  -H "anthropic-beta: code-execution-2025-08-25,skills-2025-10-02" \
  -H "content-type: application/json" \
  -d '{
    "model": "claude-sonnet-5",
    "max_tokens": 4096,
    "container": {
      "skills": [
        { "type": "custom", "skill_id": "skill_01AbCdEfGhIjKlMnOpQrStUv", "version": "latest" }
      ]
    },
    "tools": [{ "type": "code_execution_20260521", "name": "code_execution" }],
    "messages": [
      { "role": "user", "content": "Write a tailored resume for this posting: ..." }
    ]
  }'
```

Claude decides on its own whether the skill is relevant, same as the desktop
app. Any file it produces comes back as a file ID in the response — download
it with the Files API (`GET /v1/files/{id}/content`, beta header
`files-api-2025-04-14`).

### Note on the Apps Script pipeline

`app-scripts/job-tailoring.gs` does **not** use this skill — it calls the
plain Messages API (no `container`, no code-execution tool) with its own
hand-tuned `BRIEF_SYSTEM` / `LETTER_SYSTEM` / `OUTREACH_SYSTEM` prompts,
built specifically around the Job Score sheet's data and the user's
Director-of-Engineering-to-Data-Engineering positioning. Wiring this skill
into that pipeline would mean switching those calls to the container +
code-execution shape and pulling output back out as files via the Files API,
instead of the plain text the sheet-writing code expects today — a real
architecture change, not a drop-in swap. Until/unless that's wanted, treat
this skill as the interactive (desktop/API) counterpart to the pipeline, not
a replacement for it — and pull formatting rules you like (ATS rules, CAR
bullets, section colors) into the `.gs` system prompts by hand if you want
the two to match more closely.
