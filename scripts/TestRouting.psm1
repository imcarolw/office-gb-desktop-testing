Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-OfficeTestCatalog {
    param([Parameter(Mandatory)][string]$SolutionRoot)

    $catalog = @(Get-Content -LiteralPath (Join-Path $SolutionRoot 'data\test-catalog.json') -Raw |
        ConvertFrom-Json -ErrorAction Stop)
    if ($catalog.Count -eq 0) { throw 'Test catalog is empty.' }
    $ids = @{}
    $methods = @{}
    $projects = @{}
    foreach ($case in $catalog) {
        if ($case.application -notin @('Word', 'Excel', 'PowerPoint') -or
            $case.caseId -isnot [long] -and $case.caseId -isnot [int] -or $case.caseId -le 0 -or
            $case.expectedSteps -isnot [long] -and $case.expectedSteps -isnot [int] -or $case.expectedSteps -le 0 -or
            [string]::IsNullOrWhiteSpace($case.title) -or
            $case.testName -notmatch '^Case\d+_[A-Za-z0-9_]+$' -or
            -not $case.testName.StartsWith("Case$($case.caseId)_", [StringComparison]::Ordinal) -or
            @($case.aliases).Count -eq 0 -or @($case.aliases | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
            throw "Invalid test catalog entry: $($case | ConvertTo-Json -Compress)"
        }
        if ($ids.ContainsKey($case.caseId) -or $methods.ContainsKey($case.testName)) {
            throw "Duplicate case ID or test name in catalog: $($case.caseId)."
        }
        if ([IO.Path]::IsPathRooted($case.project) -or $case.project -notmatch '^tests\\[^\\]+\\[^\\]+\.csproj$') {
            throw "Catalog project must be a solution-relative test project: $($case.project)."
        }
        $projectPath = Join-Path $SolutionRoot $case.project
        if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf)) {
            throw "Catalog project does not exist: $projectPath"
        }
        if ($projects.ContainsKey($case.application) -and $projects[$case.application] -ne $case.project) {
            throw "Application $($case.application) has inconsistent projects."
        }
        if (@($projects.Keys | Where-Object { $_ -ne $case.application -and $projects[$_] -eq $case.project }).Count -gt 0) {
            throw "Catalog project is shared by different applications: $($case.project)."
        }
        $ids[$case.caseId] = $true
        $methods[$case.testName] = $true
        $projects[$case.application] = $case.project
    }
    return $catalog
}

function Select-OfficeTestCases {
    param(
        [Parameter(Mandatory)][object[]]$Catalog,
        [string[]]$Application = @('All'),
        [int[]]$CaseId = @()
    )

    if ($Application.Count -eq 0 -or @($Application | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
        throw 'Application selection must not be empty.'
    }
    if ('All' -in $Application -and $Application.Count -ne 1) { throw 'All cannot be combined with named applications.' }
    foreach ($app in $Application) {
        if ($app -ne 'All' -and $app -notin $Catalog.application) { throw "Unknown application: $app." }
    }
    if (0 -in $CaseId) {
        if ($CaseId.Count -ne 1) { throw 'CaseId 0 cannot be combined with other IDs.' }
        $CaseId = @()
    }
    foreach ($id in $CaseId) {
        $entry = @($Catalog | Where-Object caseId -eq $id)
        if ($entry.Count -ne 1) { throw "Unknown case ID: $id." }
        if ('All' -notin $Application -and $entry[0].application -notin $Application) {
            throw "Case $id belongs to $($entry[0].application), outside the selected applications."
        }
    }
    $selected = @($Catalog | Where-Object {
        ('All' -in $Application -or $_.application -in $Application) -and
        ($CaseId.Count -eq 0 -or $_.caseId -in $CaseId)
    })
    if ($selected.Count -eq 0) { throw 'No tests match the selection.' }
    return $selected
}

function New-OfficeTestPlan {
    param(
        [Parameter(Mandatory)][object[]]$Cases,
        [bool]$Interactive
    )

    $projects = @(
        foreach ($project in @($Cases.project | Select-Object -Unique)) {
            $group = @($Cases | Where-Object project -eq $project)
            [pscustomobject]@{
                Application = $group[0].application
                Project = $project
                CaseIds = @($group.caseId)
                Filter = if ($Interactive) {
                    'TestCategory=Interactive&(' + (($group | ForEach-Object { "TestCategory=Case$($_.caseId)" }) -join '|') + ')'
                } else { 'TestCategory!=Interactive' }
            }
        }
    )
    [pscustomobject]@{
        Mode = if ($Interactive) { 'Interactive' } else { 'Infrastructure' }
        Cases = @($Cases | Where-Object { $Interactive })
        Projects = $projects
    }
}

Export-ModuleMember -Function Read-OfficeTestCatalog, Select-OfficeTestCases, New-OfficeTestPlan
