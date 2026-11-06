[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(ValueFromRemainingArguments)]
    [string[]]$BridgeArguments
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$candidates = @()
# Self-contained skill install: the bridge sits next to this wrapper.
$candidates += Join-Path $PSScriptRoot 'mathtype-word.ps1'
if (-not [string]::IsNullOrWhiteSpace($env:MATHTYPE_WORD_PLUGIN_ROOT)) {
    $candidates += Join-Path $env:MATHTYPE_WORD_PLUGIN_ROOT 'scripts\mathtype-word.ps1'
}
# Plugin repository layout: skills\mathtype-for-word\scripts\invoke-bridge.ps1
$candidates += [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\scripts\mathtype-word.ps1'))
$bridge = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if ($null -eq $bridge) {
    throw 'The mathtype-for-word bridge was not found next to this wrapper. Reinstall the skill or set MATHTYPE_WORD_PLUGIN_ROOT.'
}
$editingModule = Join-Path (Split-Path -Parent $bridge) 'mathtype-word-editing.ps1'
if (-not (Test-Path -LiteralPath $editingModule -PathType Leaf)) {
    throw "The bridge is incomplete: $editingModule is missing."
}
& pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge @BridgeArguments
exit $LASTEXITCODE
