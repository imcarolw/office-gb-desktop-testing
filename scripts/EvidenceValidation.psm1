Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-OfficeRunEvidence {
    param([Parameter(Mandatory)][string]$RunDirectory)

    $RunDirectory = [IO.Path]::GetFullPath($RunDirectory)
    $errors = [Collections.Generic.List[string]]::new()
    $summary = [Collections.Generic.List[object]]::new()
    try {
        $manifest = Get-Content -LiteralPath (Join-Path $RunDirectory 'run-manifest.json') -Raw |
            ConvertFrom-Json -ErrorAction Stop
        if ($manifest.State -ne 'Completed' -or $manifest.RunId -ne (Split-Path $RunDirectory -Leaf) -or
            @($manifest.Cases).Count -eq 0) {
            throw 'Manifest is incomplete or does not identify this run.'
        }
        $start = [datetime]$manifest.StartedUtc
        $end = [datetime]$manifest.CompletedUtc
        if ($end -lt $start) { throw 'Manifest timestamps are reversed.' }
        if (@($manifest.Cases.caseId | Select-Object -Unique).Count -ne @($manifest.Cases).Count) {
            throw 'Manifest contains duplicate cases.'
        }
        foreach ($case in $manifest.Cases) {
            if ($case.application -cnotin @('Word', 'Excel', 'PowerPoint') -or [int]$case.caseId -le 0 -or
                [int]$case.expectedSteps -le 0 -or $case.project -notmatch '^tests\\[^\\]+\\[^\\]+\.csproj$' -or
                $case.testName -notmatch "^Case$($case.caseId)_[A-Za-z0-9_]+$") {
                throw 'Invalid case identity in manifest.'
            }
        }
        if (@($manifest.Executions).Count -ne @($manifest.Cases.project | Select-Object -Unique).Count) {
            throw 'Execution count does not match the selected projects.'
        }
        foreach ($cleanupError in $manifest.CleanupErrors) { $errors.Add("Cleanup: $cleanupError") }
        foreach ($project in @($manifest.Cases.project | Select-Object -Unique)) {
            $expected = @($manifest.Cases | Where-Object project -eq $project)
            $executions = @($manifest.Executions | Where-Object Project -eq $project)
            if ($executions.Count -ne 1) { throw "Expected one execution for project $project." }
            $execution = $executions[0]
            if ($execution.ExitCode -ne 0 -or $execution.Error) {
                $errors.Add("Project $project failed: exit $($execution.ExitCode); $($execution.Error)")
            }
            $trxPath = Join-Path $RunDirectory ("test-results\{0}\results.trx" -f $expected[0].application)
            [xml]$trx = Get-Content -LiteralPath $trxPath -Raw -ErrorAction Stop
            $results = @($trx.SelectNodes("//*[local-name()='UnitTestResult']"))
            $actualNames = @($results | ForEach-Object { $_.testName } | Sort-Object)
            $expectedNames = @($expected.testName | Sort-Object)
            if ($results.Count -ne $expected.Count -or
                ($actualNames -join "`n") -cne ($expectedNames -join "`n") -or
                @($results | Where-Object outcome -ne 'Passed').Count -gt 0) {
                $errors.Add("TRX does not contain exactly the selected passing tests for $project.")
            }
        }
        Add-Type -AssemblyName System.Drawing
        foreach ($case in $manifest.Cases) {
            $caseErrors = [Collections.Generic.List[string]]::new()
            $count = 0
            $stepCount = 0
            $status = 'MissingReport'
            $path = Join-Path $RunDirectory "$($case.application)\$($case.caseId)\result.json"
            try {
                $report = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -ErrorAction Stop
                $status = $report.Status
                $stepCount = @($report.Steps).Count
                if ($report.Application -cne $case.application -or $report.CaseId -ne $case.caseId -or
                    $report.ExpectedStepCount -ne $case.expectedSteps -or
                    $stepCount -ne $case.expectedSteps -or $status -ne 'NeedsVisualReview' -or
                    $report.VisualStatus -ne 'ReviewRequired' -or $report.Error) {
                    throw 'Case identity, status, or step coverage does not match the requested test.'
                }
                if ([datetime]$report.StartedUtc -lt $start -or -not $report.CompletedUtc -or
                    [datetime]$report.CompletedUtc -gt $end) { throw 'Case timestamps are outside this run.' }
                for ($i = 0; $i -lt $stepCount; $i++) {
                    $step = $report.Steps[$i]
                    if ($step.Number -ne $i + 1 -or $step.Status -ne 'AutomatedChecksPassed' -or
                        $step.Error -or $step.ScreenshotError -or -not $step.CompletedUtc) {
                        throw "Step $($i + 1) is incomplete or failed."
                    }
                    $stepStart = [datetime]$step.StartedUtc
                    $stepEnd = [datetime]$step.CompletedUtc
                    if ($stepStart -lt [datetime]$report.StartedUtc -or $stepEnd -lt $stepStart -or
                        $stepEnd -gt [datetime]$report.CompletedUtc) { throw 'Invalid step timestamps.' }
                    $names = @($step.Screenshots)
                    $prefix = 'step-{0:D2}-' -f ($i + 1)
                    if ($names.Count -eq 0 -or @($names | Where-Object { $_ -match "^$prefix\d+-after\.png$" }).Count -ne 1 -or
                        @($names | Select-Object -Unique).Count -ne $names.Count) {
                        throw "Step $($i + 1) needs unique screenshots and exactly one after-step checkpoint."
                    }
                    foreach ($name in $names) {
                        if ($name -cne [IO.Path]::GetFileName($name) -or $name -notmatch "^$prefix\d+-[A-Za-z0-9-]+\.png$") {
                            throw "Invalid screenshot reference: $name"
                        }
                        $imagePath = Join-Path (Split-Path $path) $name
                        $image = [Drawing.Image]::FromFile($imagePath)
                        try {
                            if ($image.Width -lt 100 -or $image.Height -lt 100 -or
                                $image.RawFormat.Guid -ne [Drawing.Imaging.ImageFormat]::Png.Guid) {
                                throw "Not a usable PNG screenshot: $name"
                            }
                        } finally { $image.Dispose() }
                        $written = (Get-Item -LiteralPath $imagePath).LastWriteTimeUtc
                        if ($written -lt $stepStart -or $written -gt $stepEnd) {
                            throw "Screenshot was not written during its recorded step: $name"
                        }
                        $count++
                    }
                }
            } catch {
                $caseErrors.Add($_.Exception.Message)
                $errors.Add("Case $($case.caseId): $($_.Exception.Message)")
            }
            $summary.Add([pscustomobject]@{
                Application = $case.application; CaseId = $case.caseId; Status = $status
                Steps = $stepCount; Screenshots = $count; Report = $path
                EvidenceValid = $caseErrors.Count -eq 0; Errors = @($caseErrors)
            })
        }
        $expectedReports = @($manifest.Cases | ForEach-Object {
            [IO.Path]::GetFullPath((Join-Path $RunDirectory "$($_.application)\$($_.caseId)\result.json"))
        })
        $reports = @(
            Get-ChildItem -LiteralPath $RunDirectory -Filter result.json -File
            Get-ChildItem -LiteralPath $RunDirectory -Directory |
                Where-Object Name -ne 'test-results' |
                ForEach-Object { Get-ChildItem -LiteralPath $_.FullName -Filter result.json -Recurse -File }
        )
        $unexpected = @($reports | Where-Object { $_.FullName -notin $expectedReports })
        if ($unexpected.Count -gt 0) { $errors.Add('Unexpected case reports exist in this run.') }
    } catch {
        $errors.Add("Run validation: $($_.Exception.Message)")
    }
    [pscustomobject]@{
        RunDirectory = $RunDirectory
        Passed = $errors.Count -eq 0
        VisualStatus = 'ReviewRequired'
        Cases = @($summary)
        Errors = @($errors)
    }
}

function Write-OfficeRunSummary {
    param([Parameter(Mandatory)]$Validation)

    $directory = $Validation.RunDirectory
    $lines = @(
        '# Office GB test results'
        ''
        "**Automated result: $(if ($Validation.Passed) { 'Passed' } else { 'Failed' })**"
        ''
        'Visual glyph review is still required. Automated checks do not certify rendering.'
        ''
        '| Application | Case | Automated checks | Steps | Validated screenshots | Evidence |'
        '| --- | --- | --- | --- | --- | --- |'
        foreach ($case in $Validation.Cases) {
            $result = if ($case.EvidenceValid -and $case.Status -eq 'NeedsVisualReview') { 'Passed' } else { 'Failed' }
            $relative = [IO.Path]::GetRelativePath($directory, $case.Report).Replace('\', '/')
            $folder = $relative.Substring(0, $relative.LastIndexOf('/') + 1)
            "| $($case.Application) | $($case.CaseId) | $result | $($case.Steps) | $($case.Screenshots) | [Report]($relative) / [Screenshots]($folder) |"
        }
        ''
        '[Technical validation](evidence-validation.json) | [Machine-readable summary](run-summary.json)'
        if (@($Validation.Errors).Count -gt 0) {
            ''
            '## Failures'
            ''
            '```text'
            $Validation.Errors
            '```'
        }
    )
    $path = Join-Path $directory 'run-summary.md'
    $lines | Set-Content -LiteralPath $path -Encoding utf8
    return $path
}

Export-ModuleMember -Function Test-OfficeRunEvidence, Write-OfficeRunSummary
