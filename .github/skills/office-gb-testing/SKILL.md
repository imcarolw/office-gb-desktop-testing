---
name: office-gb-testing
description: 'Run or list the implemented Office GB18030 desktop tests for Word, Excel, and PowerPoint. Use when asked to run all Excel tests, Word tests, PowerPoint tests, or an Office case ID or scenario such as Save As, text boxes, default fonts, or Find and Replace. Select tests from the shared catalog, execute sequentially, and report screenshot evidence.'
argument-hint: 'Application, case ID or scenario'
user-invocable: true
---

# Office GB desktop testing

## Locate the suite

This skill is installed inside the standalone automation workspace. Resolve the
suite root relative to this file: three parents above this skill's directory.
The runner is `Run-Tests.ps1` directly under that root. Use absolute resolved
paths for execution, never a fixed drive, username, or current-working-directory
assumption. Read [the suite guide](../../../README.md) as needed.
Use PowerShell 7 (`pwsh`).

## Select tests from the catalog

1. Read [the shared catalog](../../../data/test-catalog.json)
   or invoke the suite's `Run-Tests.ps1 -ListTests`. Do not maintain another case
   list here. "All Excel tests" means all implemented Excel catalog entries,
   not every possible Excel scenario. Use only this repository's runner.
2. Resolve application names case-insensitively. "Power point" and "PPT" mean
   PowerPoint. Match scenario descriptions using the catalog's titles/aliases,
   restricting by application when specified. Exact IDs must exist in the catalog.
   For multiple applications or IDs, pass PowerShell arrays as typed arguments.
   Do not interpolate the user's free-form request into an executable command.
3. "Run Word tests" selects all Word cases; "run all Office tests" selects All.
   A specific scenario selects only the matching case(s). For "any PowerPoint
   test", conflicting IDs/applications, an ambiguous scenario, or a request
   outside the catalog, ask for clarification instead of guessing.
4. Announce a short selection message with application(s), case IDs, and scenarios.
   Do not expose project paths, NUnit filters, or catalog JSON in the normal demo.
   The runner validates selection; a separate preview is optional for troubleshooting,
   not a prerequisite for execution.

Use these selectors with the resolved runner path:

```powershell
& $runner -Interactive -Application Excel
& $runner -Interactive -Application Word,Excel
& $runner -Interactive -CaseId <catalog-case-id>
& $runner -Interactive -CaseId <id1>,<id2>
```

Do not run infrastructure tests before the demo. `-Preview`, `-ListTests`, the
standalone validator, and infrastructure-only execution are developer tools;
use them only when requested or needed for troubleshooting.

## Desktop safety and execution

- A request to run tests authorizes the requested selection, not closing user
  applications, deleting evidence, changing tests/settings, or installing software.
- Before interactive execution, confirm an unlocked local Windows desktop is
  ready and the user can remain hands-off. If already confirmed for the current
  run, do not ask again. A remote/headless agent cannot run these tests by itself;
  execution must occur in the interactive Windows session with Office installed.
- The runner refuses pre-existing Word/Excel/PowerPoint sessions. If it reports
  any, ask the user to close them; never kill or attach to a user session.
  Keep the harness's exclusive desktop lock and opt-in gate intact.
- Invoke the existing runner, not handwritten `dotnet` filters or COM UI
  substitutes. Each invocation runs the selected projects sequentially once.
  When the user asks to rerun, invoke the runner again after the previous run
  finishes; it creates a fresh evidence folder without overwriting earlier results.
- Monitor the same running command until completion; never launch overlapping
  UI runs. Do not detach the process or assume successful launch means completion.
- Do not change tests, weaken assertions, silently retry failures, or substitute
  unsupported cases. Unexpected leftover Office processes are reported as a
  failure; do not close them without approval.
- Do not clean up historical evidence unless the user explicitly asks.

## Validate and report

1. Capture the exact `Results summary` path printed by this execution and read
   that `run-summary.md`. Link to this single readable summary, which shows
   application, case, automated outcome, steps, screenshot counts, and evidence.
   Do not choose an old or merely newest folder.
2. Evidence validation is automatic; do not run the standalone validator again
   during the normal demo. A nonzero exit, missing summary, failed validation,
   skipped/blocked case, or unfinished run must not be reported as successful.
   On failure, inspect the linked technical validation and relevant case reports.
3. Give a concise outcome and the summary link. Keep technical reports behind
   that summary unless troubleshooting. A successful manual rerun does not erase
   or change an earlier failure.
4. State that `NeedsVisualReview` means automated checks passed, not visual GB
   certification. Correct Unicode readback and font metadata do not prove glyph
   rendering. Never mark visual review complete unless explicitly performed
   against agreed acceptance criteria.

If skill discovery does not occur, have the user open the repository root in VS
Code, check that agent skills are enabled/available, and invoke
`/office-gb-testing` explicitly. Frontmatter guides selection; it is not a
guarantee of model routing or an OS permission boundary.
