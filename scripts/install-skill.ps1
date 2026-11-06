[CmdletBinding(DefaultParameterSetName = 'Roots')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Explicit')]
    [string[]] $SkillRoot,

    [Parameter(Mandatory = $true, ParameterSetName = 'Roots')]
    [switch] $DefaultRoots,

    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$source = Join-Path $repoRoot 'mathtype-for-word'
if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md') -PathType Leaf)) {
    throw "Invalid repository layout: $source\SKILL.md is missing."
}

$roots = if ($DefaultRoots) {
    @(
        (Join-Path $env:USERPROFILE '.codex\skills'),
        (Join-Path $env:USERPROFILE '.agents\skills'),
        (Join-Path $env:USERPROFILE '.dsh\skills')
    )
}
else {
    $SkillRoot
}

foreach ($root in $roots) {
    if ([string]::IsNullOrWhiteSpace($root)) { continue }
    $resolvedRoot = [IO.Path]::GetFullPath($root)
    New-Item -ItemType Directory -Force -Path $resolvedRoot | Out-Null
    $target = Join-Path $resolvedRoot 'mathtype-for-word'

    if (Test-Path -LiteralPath $target) {
        if (-not $Force) {
            throw "Destination exists; pass -Force to replace it: $target"
        }
        $item = Get-Item -LiteralPath $target -Force
        if ($item.LinkType -in @('Junction', 'SymbolicLink')) {
            Remove-Item -LiteralPath $target -Force
        }
        else {
            Remove-Item -LiteralPath $target -Recurse -Force
        }
    }

    Copy-Item -LiteralPath $source -Destination $target -Recurse -Force
    if (-not (Test-Path -LiteralPath (Join-Path $target 'SKILL.md') -PathType Leaf)) {
        throw "Copy failed: $target\SKILL.md is missing."
    }
    Write-Host "Installed mathtype-for-word -> $target" -ForegroundColor Green
}
