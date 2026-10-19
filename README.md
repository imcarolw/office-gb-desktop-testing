# Office GB character desktop automation

One .NET 8 / NUnit / FlaUI UIA3 solution, with separate Word, Excel, and
PowerPoint test projects and a shared desktop/evidence library. This standalone
prototype implements six scenarios with a screenshot after every numbered step.
No external demo project, source workbook, account, or cloud storage is required
by the test harness. Office must already be installed, licensed, and ready to use.

## Get started on another Windows PC

Install Git, PowerShell 7, a .NET SDK 9 or later plus the .NET 8 runtime, and
desktop Office. Then clone this repository and open its root folder:

```powershell
git clone <repository-url> office-gb-desktop-testing
Set-Location .\office-gb-desktop-testing
dotnet build .\OfficeGbAutomation.slnx
```

The build restores public NuGet dependencies. No machine-specific paths need to
be edited for a standard Office installation. Use the demo quick start below;
VS Code and an agent supporting workspace skills are optional for command-line
execution. For a developer-only check that does not launch Office, run
`.\Run-Tests.ps1`.

## Demo quick start

1. Open the cloned repository root in VS Code with agent
   skills enabled. Use an unlocked Windows desktop with Office installed and
   close Word, Excel, and PowerPoint before starting.
2. Ask **"Run all Excel tests"**, **"Run Word Save As"**, or
   **"Run PowerPoint Find and Replace"**. Alternatively invoke `/office-gb-testing`.
3. The skill announces the selected cases and confirms desktop readiness once.
   Leave the mouse and keyboard untouched while the selected tests run.
4. Open the linked **results summary**: one table with application, case,
   automated outcome, steps, screenshot counts, and evidence links.

The demo flow is **request -> select -> run Office -> show results**.
No separate preview, infrastructure test run, or manual evidence-validation
command is required. Evidence validation still runs automatically, and every
numbered step still has a screenshot. `NeedsVisualReview` means automated checks
passed, not that glyph rendering has been visually certified.

For direct execution from PowerShell 7 in this solution folder:

```powershell
.\Run-Tests.ps1 -Interactive -Application Excel
```

Each invocation runs once and creates a new evidence folder. Ask to run again
for a manual rerun; earlier results are preserved. The remaining sections cover
developer setup, detailed scope, and troubleshooting.

## Case mapping

The included catalog and tests define the six implemented cases. Every numbered
step has a mandatory screenshot; no separate spreadsheet is needed to run them.

| Sheet | Case | Scenario | Steps |
| --- | --- | --- | --- |
| Word for Win32 | 5842 | File -> New template search | 4 |
| Word for Win32 | 5844 | GB text and filename through File -> Save As -> Browse -> Desktop | 6 |
| Excel for Win32 | 6841 | GB workbook/sheet names in Options -> Advanced | 8 |
| Excel for Win32 | 5881 | Horizontal and vertical text boxes | 4 |
| PowerPoint for Win32 | 5921 | Home -> Find and Replace | 12 |
| PowerPoint for Win32 | 6295 | Title/subtitle text and default Deng Xian font | 4 |

The demo uses local Excel text-box testing instead of a cloud-storage scenario,
and PowerPoint Find and Replace instead of legacy template search. Only the six
listed cases are implemented; other case IDs are not supported.

## Prerequisites and safety

- Windows, PowerShell 7 (`pwsh`), .NET 8 runtime and SDK 9 or later (the solution uses `.slnx`).
- Desktop Word, Excel, and PowerPoint; no first-run/sign-in/recovery dialogs.
- An unlocked interactive desktop, at the same elevation as Office.
- Close **all** Word, Excel, and PowerPoint processes before running. Do not use
  the mouse/keyboard while automation runs. A shared lock prevents concurrent
  runs using the same artifact root; always keep one common root.
- UI labels support English and selected Simplified Chinese Office labels.
  Other localizations need matching selectors; missing controls fail explicitly.
- Excel 5881 needs the horizontal/vertical text-box menu commands specified in
  the workbook. Builds without the vertical command fail rather than omit it.
- PowerPoint 6295 checks the default East Asian font; newer default themes may
  legitimately fail the legacy Deng Xian expectation. The test never forces a
  font to manufacture a pass.
- The known "Cloud Create Onboarding" dialog is closed through its Close button
  only when present, without opting into cloud saving or changing preferences.
- After creating or opening a document, the runner also dismisses visible
  OfficePLUS/Copilot task panes identified by their exact Office pane markers
  inside the test-owned application. Already-closed panes require no action.
  Pane content, headers, and Close buttons must have nonempty bounds intersecting
  the Office window and positive visibility evidence. An unsupported UI Automation
  offscreen property is not treated as visible; a matching, visible native child
  window owned by the test process can supply that evidence instead. Hidden panes
  and stale/unknown entries are not closed merely because their names exist.
  It invokes only the Close pane button in that pane's matching named native
  header, outside embedded web content. It never uses
  the application's title-bar Close button, disables add-ins, or edits Office
  preferences. Other unexpected dialogs remain explicit failures.
- Startup waits for the actual document (native Word/Excel editing surface or
  a PowerPoint presentation containing a slide) and the visible Home/File ribbon.
  Only known transient UI Automation/Office-busy COM errors are retried within
  the configured wait timeout; other errors propagate immediately.
- If an optional pane cannot be closed safely, keeps reopening, or leaves Office
  controls inaccessible, the run fails with diagnostics. Close the pane manually
  and investigate add-in compatibility; closing a pane cannot guarantee recovery
  from an add-in that disrupts Office accessibility.
- Desktop screenshots may include pixels behind overlapping Office windows.
  Use a clean test desktop without sensitive content.

Configure Office paths, timeouts, and artifact root in
[settings/office-test-settings.json](settings/office-test-settings.json).
By default, `officeDirectory` is empty: each executable is discovered through
Windows App Paths registration (per-user/machine, 64-bit/32-bit registry views).
Set `officeDirectory` only for a custom installation; a missing configured
executable fails explicitly rather than silently switching installations.
`artifactRoot` defaults to `artifacts`, relative to the solution root, not the
terminal's working directory. An explicit Office directory can also be relative
to that root; absolute overrides remain supported.

The runner obtains the root from its own location and passes `OFFICE_TEST_ROOT`
to the tests. Direct `dotnet test` locates the solution above its output folder.
For separately deployed test binaries, set `OFFICE_TEST_ROOT` to the desired
existing root directory. Desktop and temporary directories are obtained from
Windows APIs, never from a fixed user profile path.
Samples in [data/gb-samples.json](data/gb-samples.json) include Chinese and
supplementary-plane characters. These are representative samples, not an
exhaustive GB18030 certification corpus.

## Developer commands

From PowerShell in the solution folder (wherever you copied it):

```powershell
dotnet build .\OfficeGbAutomation.slnx
.\Run-Tests.ps1                         # Infrastructure tests only; no Office UI
.\Run-Tests.ps1 -Interactive            # All six cases, applications sequentially
.\Run-Tests.ps1 -Interactive -Application Excel
.\Run-Tests.ps1 -Interactive -CaseId 5844
.\Run-Tests.ps1 -Interactive -Application Word,Excel
.\Run-Tests.ps1 -Interactive -CaseId 5844,5921
.\Run-Tests.ps1 -Application Excel -ListTests
.\Run-Tests.ps1 -Interactive -Application Excel -Preview
```

Direct `dotnet test` skips UI cases unless `RUN_OFFICE_UI_TESTS=1`. Do not run the
solution's interactive tests in parallel; use the runner above.

The runner works from any working directory; invoke its resolved path with `&`.
The default remains infrastructure-only and clears inherited UI opt-in while
running. Infrastructure checks run in each selected application's project,
including catalog/NUnit consistency checks. Case IDs require `-Interactive`,
except in `-ListTests`; legacy `-CaseId 0` still means no ID restriction.
Combining `All` with named applications, unknown IDs, and application/ID
mismatches fails explicitly. `-Preview` and `-ListTests` return JSON without
building, launching Office, or creating evidence.

The shared [test catalog](data/test-catalog.json) is the inventory for routing,
scenario aliases, expected steps, projects, and NUnit methods. Add future
implemented cases there alongside their NUnit tests; infrastructure checks
reject inconsistent method inventories and runtime checks verify step counts.
Application/case arrays are unions, deduplicated by the catalog; application
and case restrictions must agree.

## Start through the workspace skill

Open the **repository root** as the VS Code workspace. The workspace skill is
[office-gb-testing](.github/skills/office-gb-testing/SKILL.md). With agent skills
available/enabled, use `/office-gb-testing` or requests such as:

- "Run all Excel tests."
- "Run Word Save As."
- "Run PowerPoint Find and Replace."
- "Run Word and Excel tests."
- "List the available PowerPoint tests."

The skill reads the catalog, resolves and announces the selection, confirms the
hands-off desktop is ready once, then invokes the existing runner. A preview is
optional for troubleshooting, not a required demo step. "All" means all
**implemented catalog cases**, not every possible Office scenario. Ambiguous requests
such as "any PowerPoint test" require clarification. Natural-language routing is
agent guidance, not a deterministic language parser; actual IDs, filters, safety,
and evidence are validated in the runner. A headless/cloud agent cannot execute
this desktop suite without an interactive Windows test host.

## Manual reruns and validation

Each invocation runs the selected cases once. To rerun, invoke the same command
again or ask the skill to run the tests again after the current run finishes.
Each invocation creates a unique run ID with its own manifest, TRX, screenshots,
reports, and validation result, so earlier results are never overwritten.
There is no automatic retry, repeat parameter, or batch summary. Remaining
Office processes are reported as a failure and are not killed by the runner.
An interrupted manifest remains `Running` and cannot validate as a complete run.

The primary demo report is `run-summary.md`, generated automatically alongside
the technical JSON/TRX files. It links to each case report and screenshot folder;
failures appear explicitly with links to technical validation. Infrastructure
tests are developer checks and are not run before or during interactive demos.

For developer troubleshooting only, recheck one recorded run:

```powershell
.\Validate-Evidence.ps1 -RunDirectory <path-to-the-exact-run-folder>
```

This emits structured JSON and fails with a nonzero exit code for invalid
evidence. It checks the selected-case manifest against the exact TRX test set
and passing outcomes, case identities/statuses, sequential expected steps,
per-step after images, filename containment, decoded PNG dimensions/format,
and screenshot write times within their recorded steps. Missing, corrupt,
stale, incomplete, blocked, skipped, or extra case evidence does not pass.
The validator verifies consistency, not cryptographic authenticity.

Infrastructure tests include
[no-UI workflow checks](scripts/Test-Workflow.ps1), run through NUnit. They use
an isolated temporary fixture with four focused checks: application/scenario
routing, unknown/conflicting selection rejection, evidence/summary outcomes
(including failed cases and missing screenshots), and unique folders preserving
earlier manual runs. A small `dotnet` stand-in exercises the no-UI runner's
run-ID lifecycle; no Office process or UI execution is simulated.
Those temporary synthetic images are validator test fixtures,
not real Office run evidence; the temporary fixture is removed after testing.

## Evidence and acceptance

`artifacts\<unique-run-id>\<application>\<case-id>\result.json` records each
numbered step, UTC timestamps, action, status, screenshots, and explicit errors.
Every completed step captures `step-NN-NN-after.png`. Multi-action steps also
capture intermediate checkpoints. Failure attempts capture the current UI,
write an additional UI tree diagnostic, and stop the case. Screenshot failures
are failures, never ignored. No screenshots are fabricated if launch fails.
Unexecuted later steps have no evidence and are not claimed as completed.
The runner groups selected cases under one run ID and writes a readable
`run-summary.md` for the demo, plus `run-summary.json`
with case status, completed steps, screenshot counts, report paths, and evidence
validity. `run-manifest.json` pins the selected catalog entries, timestamps, and
project exit codes. `evidence-validation.json` records explicit validation errors. Missing,
failed, or blocked reports cause the runner to fail, even if NUnit marks an
unsupported case as inconclusive.

NUnit passing means **automated checks passed**, not full visual certification.
Successful case JSON has `Status: NeedsVisualReview` and
`VisualStatus: ReviewRequired`. Review every screenshot for missing-glyph boxes,
question marks, clipping, font fallback, and incorrect glyphs before accepting
the case. No OCR or unapproved image baseline is used as a proxy for rendering.

Interactive NUnit/TRX outcomes are in each run's `test-results` subfolder.
Infrastructure TRX outcomes are under `TestResults`. Reports and screenshots are
persisted after each step, including on assertion failures.

Word 5844 saves a unique file on the actual Desktop, and copies it into the case
artifact directory. The Desktop file is intentionally retained as evidence;
delete only that named generated document when no longer needed.

Excel 6841 prepares a local GB-named `.xlsx` fixture through COM before UI launch.
The test opens it through File -> Open -> Browse. The source step 4's "Click Save"
is interpreted as Options -> Save, where AutoRecover exceptions are located in
current Excel. Step 5 opens Advanced to inspect worksheet names; step 6 navigates
there again. Its image/workbook/sheet dropdown checkpoints are recorded separately.

PowerPoint 5921 supports the current combined Find and Replace pane as well as
legacy named fields. On the modern pane, its two unlabeled edit fields are
identified inside that specific pane in vertical order; Enter in the Find field
performs Find Next. Found selections and final replaced text are read back
exactly. Step 8 uses the already-open modern pane's replacement field rather
than closing and reopening it; this follows the source case's direct transition
from Find to Replace. The replacement is the configured GB-character filename stem.

PowerPoint 6295 uses native slide-to-screen coordinate conversion to click
placeholders when their labels are absent from UIA. No COM selection or text
assignment is used. Default East Asian theme tokens are resolved against the
exported font scheme's `Hans` (Simplified Chinese) entries; empty generic East
Asian theme names are not treated as actual fonts. Latin Aptos in the ribbon is
not the same as the theme's Han font. Font assertions do not prove which fallback
font rendered every supplementary-plane glyph; screenshot review remains required.

COM is used only for fixture preparation and deterministic readback (text,
sheet names, orientations, fonts, and placeholder geometry), including font-scheme
export for theme inspection. Input, navigation, drawing, opening, replacement,
and the Word save operation use the real UI. DPI-aware screen capture prevents
cropping on scaled displays. Cleanup targets only the launched process.

## Repository contents and local evidence

The repository contains only test source, project files, generic GB-character
samples, configuration, runner/validation scripts, developer checks, the workspace
skill, and documentation. The developer checks are optional for a live demo.

Build outputs, local editor settings, screenshots, saved Office documents,
diagnostic UI dumps, and test reports are not distributed. Each user generates
fresh evidence locally under the configured artifact root. Reports can contain
usernames, machine paths, and Office UI content; do not commit or upload them
without reviewing them for private information.

Office versions, language, fonts, and desktop state can affect UI results.
Always use the fresh summary from your own run; automated success is not a
guarantee of visual glyph correctness.

Startup retries and pane dismissals are recorded in the console and in the case
report's `Observations.StartupHandling`. Failure screenshots and `failure-ui.txt`
remain local diagnostics. The automated startup tests cover pane absence,
closure, missing Close controls, failed closure, exact marker selection, and
transient-error handling. Testing against a real OfficePLUS installation is
still required to verify compatibility with that add-in's particular version.
