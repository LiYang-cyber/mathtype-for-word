# mathtype-word-editing.ps1
# Discovery, inventory, in-place editing, and preview support for the mathtype-for-word bridge.
# Dot-sourced by scripts/mathtype-word.ps1 so every function shares the same script scope.

# ------------------------------------------------------------------ discovery

function Get-PropertyValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-MathTypeInstallInfo {
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($env:MATHTYPE_EXECUTABLE_PATH)) { $candidates.Add($env:MATHTYPE_EXECUTABLE_PATH) }
    $command = Get-Command MathType.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $command) { $candidates.Add($command.Source) }
    $uninstallRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($root in $uninstallRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($key in @(Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            $properties = $null
            try { $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction Stop } catch { continue }
            if ($null -eq $properties) { continue }
            $displayName = [string](Get-PropertyValue -Object $properties -Name 'DisplayName' -Default '')
            if ($displayName -notlike '*MathType*') { continue }
            foreach ($raw in @(
                (Get-PropertyValue -Object $properties -Name 'DisplayIcon' -Default ''),
                (Get-PropertyValue -Object $properties -Name 'InstallLocation' -Default '')
            )) {
                if ([string]::IsNullOrWhiteSpace($raw)) { continue }
                $path = ([string]$raw).Trim().Trim('"')
                if ($path -match ',') { $path = ($path -split ',')[0].Trim() }
                if ($path -match '\.exe$') { $candidates.Add($path) }
                elseif (Test-Path -LiteralPath $path -PathType Container) { $candidates.Add((Join-Path $path 'MathType.exe')) }
            }
        }
    }
    $programFilesX86 = [Environment]::GetFolderPath('ProgramFilesX86')
    foreach ($known in @(
        (Join-Path $env:ProgramFiles 'MathType\MathType.exe'),
        (Join-Path $programFilesX86 'MathType\MathType.exe'),
        'D:\MathType\MathType.exe'
    )) { if (-not [string]::IsNullOrWhiteSpace($known)) { $candidates.Add($known) } }
    $executable = $null
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
        $executable = (Resolve-Path -LiteralPath $candidate).Path
        break
    }
    $version = $null
    if ($null -ne $executable) { $version = (Get-Item -LiteralPath $executable).VersionInfo.ProductVersion }
    return [ordered]@{
        executable = $executable
        root       = if ($null -ne $executable) { Split-Path -Parent $executable } else { $null }
        version    = $version
    }
}

function Find-MathTypeWordTemplate {
    param([string]$ExplicitPath)
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) { $candidates.Add($ExplicitPath) }
    if (-not [string]::IsNullOrWhiteSpace($env:MATHTYPE_WORD_TEMPLATE)) { $candidates.Add($env:MATHTYPE_WORD_TEMPLATE) }
    $startupRoots = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($env:APPDATA)) { $startupRoots.Add((Join-Path $env:APPDATA 'Microsoft\Word\STARTUP')) }
    foreach ($officeRoot in @($env:ProgramFiles, [Environment]::GetFolderPath('ProgramFilesX86'))) {
        if ([string]::IsNullOrWhiteSpace($officeRoot)) { continue }
        foreach ($version in @('Office16', 'Office15', 'Office14')) {
            $startupRoots.Add((Join-Path $officeRoot ('Microsoft Office\root\' + $version + '\STARTUP')))
            $startupRoots.Add((Join-Path $officeRoot ('Microsoft Office\' + $version + '\STARTUP')))
        }
    }
    if ($null -ne $script:Word) {
        try { $startupRoots.Add([string]$script:Word.StartupPath) } catch {}
    }
    foreach ($root in $startupRoots) {
        if ([string]::IsNullOrWhiteSpace($root) -or -not (Test-Path -LiteralPath $root -PathType Container)) { continue }
        foreach ($file in @(Get-ChildItem -LiteralPath $root -Filter 'MathType Commands*.dotm' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)) {
            $candidates.Add($file.FullName)
        }
    }
    $info = Get-MathTypeInstallInfo
    if ($null -ne $info.root) {
        foreach ($bits in @('64', '32')) {
            $support = Join-Path $info.root ('Office Support\' + $bits)
            if (-not (Test-Path -LiteralPath $support -PathType Container)) { continue }
            foreach ($file in @(Get-ChildItem -LiteralPath $support -Filter 'MathType Commands*.dotm' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)) {
                $candidates.Add($file.FullName)
            }
        }
    }
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).Path }
    }
    return $null
}

function Get-MathTypeVectorStyle {
    $key = 'HKCU:\Software\Design Science\DSMT7\Styles'
    $value = $null
    if (Test-Path -LiteralPath $key) {
        $properties = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
        $value = [string](Get-PropertyValue -Object $properties -Name 'Vector')
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        return [ordered]@{ configured = $false; value = $null; bold = $false; italic = $false; bold_italic = $false }
    }
    $flags = ($value -split ',')[-1].Trim().ToUpperInvariant()
    $bold = $flags.Contains('B')
    $italic = $flags.Contains('I')
    return [ordered]@{ configured = $true; value = $value; bold = $bold; italic = $italic; bold_italic = ($bold -and $italic) }
}

function Set-MathTypeVectorStyle {
    param([Parameter(Mandatory)][string]$Value, [switch]$WhatIfOnly)
    $key = 'HKCU:\Software\Design Science\DSMT7\Styles'
    if (-not (Test-Path -LiteralPath $key)) { throw 'MathType style preferences were not found under HKCU:\Software\Design Science\DSMT7\Styles.' }
    $before = [string](Get-PropertyValue -Object (Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue) -Name 'Vector')
    $backupPath = Join-Path $script:UserConfigDirectory 'vector-style-backup.json'
    $backupWritten = $false
    if (-not (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        New-Item -ItemType Directory -Path $script:UserConfigDirectory -Force | Out-Null
        [ordered]@{ captured_at = (Get-Date).ToString('o'); vector = $before } | ConvertTo-Json | Set-Content -LiteralPath $backupPath -Encoding UTF8
        $backupWritten = $true
    }
    if (-not $WhatIfOnly) { Set-ItemProperty -LiteralPath $key -Name Vector -Value $Value }
    return [ordered]@{
        before        = $before
        after         = $Value
        changed       = ($before -ne $Value)
        backup_path   = $backupPath
        backup_written = $backupWritten
    }
}

function Test-UserProfileWritable {
    $target = $env:APPDATA
    if ([string]::IsNullOrWhiteSpace($target)) { return $false }
    $probe = Join-Path $target ('.mathtype-for-word-probe-' + $script:RunToken + '.tmp')
    try {
        Set-Content -LiteralPath $probe -Value 'probe' -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch { return $false }
}

function Enable-MathTypeWordCommands {
    param([string]$TemplatePath)
    if ($null -eq $script:Word) { throw 'Word is not running.' }
    foreach ($template in @($script:Word.Templates)) {
        if ($template.Name -like 'MathType Commands*.dotm') {
            $script:TemplateLoadMode = 'loaded'
            Release-ComObject $template
            return $template.Name
        }
        Release-ComObject $template
    }
    if ($null -ne $script:RuntimeAddIn) { $script:TemplateLoadMode = 'runtime'; return $script:RuntimeAddIn.Name }
    $path = $TemplatePath
    if ([string]::IsNullOrWhiteSpace($path)) { $path = Find-MathTypeWordTemplate }
    if ([string]::IsNullOrWhiteSpace($path)) {
        throw 'MathType Commands (*.dotm) was not found. Install MathType for Word, pass -TemplatePath, or set MATHTYPE_WORD_TEMPLATE.'
    }
    try {
        foreach ($addIn in @($script:Word.AddIns)) {
            if ($addIn.Path -eq $path) {
                $addIn.Installed = $true
                $script:RuntimeAddIn = $addIn
                $script:TemplateLoadMode = 'runtime-existing'
                return $addIn.Name
            }
            Release-ComObject $addIn
        }
    }
    catch { Write-Log -Level WARN -Message ('Could not inspect existing Word add-ins: ' + $_.Exception.Message) }
    $added = $script:Word.AddIns.Add($path, $true)
    $added.Installed = $true
    $script:RuntimeAddIn = $added
    $script:TemplateLoadMode = 'runtime'
    Write-Log -Level INFO -Message ('Loaded MathType Word commands at runtime: ' + $path)
    return $added.Name
}

function Get-MathTypeWordTemplateStatus {
    param([string]$TemplatePath)
    $resolved = Find-MathTypeWordTemplate -ExplicitPath $TemplatePath
    $inStartup = $false
    if ($null -ne $resolved) {
        $startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
        if ($resolved.StartsWith($startup, [StringComparison]::OrdinalIgnoreCase)) { $inStartup = $true }
    }
    return [ordered]@{
        path           = $resolved
        discovered     = ($null -ne $resolved)
        user_startup   = $inStartup
        startup_folder = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
    }
}

# ------------------------------------------------------------------ package reading

function Get-ZipEntryText {
    param([Parameter(Mandatory)]$Zip, [Parameter(Mandatory)][string]$Name)
    $entry = $Zip.Entries | Where-Object { $_.FullName -eq $Name } | Select-Object -First 1
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Close() }
}

function Get-ZipEntryBytes {
    param([Parameter(Mandatory)]$Zip, [Parameter(Mandatory)][string]$Name)
    $entry = $Zip.Entries | Where-Object { $_.FullName -eq $Name } | Select-Object -First 1
    if ($null -eq $entry) { return $null }
    $stream = New-Object System.IO.MemoryStream
    try {
        $source = $entry.Open()
        try { $source.CopyTo($stream) } finally { $source.Close() }
        return $stream.ToArray()
    }
    finally { $stream.Dispose() }
}

function Get-MathTypeTexFromOleBytes {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    $nul = [char]0
    $latin = [System.Text.Encoding]::GetEncoding(28591).GetString($Bytes)
    $marker = 'TeX Input Language'
    $index = $latin.IndexOf($marker, [StringComparison]::Ordinal)
    if ($index -lt 0) { return $null }
    $start = $latin.IndexOf($nul, $index)
    if ($start -lt 0) { return $null }
    $start++
    $end = $latin.IndexOf($nul, $start)
    if ($end -lt 0) { return $null }
    return $latin.Substring($start, $end - $start)
}

function Get-PackageEquationInventory {
    param([Parameter(Mandatory)][string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $document = Get-ZipEntryText -Zip $zip -Name 'word/document.xml'
        $relationships = Get-ZipEntryText -Zip $zip -Name 'word/_rels/document.xml.rels'
        $targets = @{}
        foreach ($match in [regex]::Matches($relationships, '<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"')) {
            $targets[$match.Groups[1].Value] = $match.Groups[2].Value
        }
        $records = [System.Collections.Generic.List[object]]::new()
        $ordinal = 0
        $paragraphIndex = 0
        foreach ($paragraph in [regex]::Matches($document, '(?s)<w:p\b.*?</w:p>')) {
            $paragraphIndex++
            $xml = $paragraph.Value
            $objects = [regex]::Matches($xml, '<o:OLEObject\b[^>]*r:id="([^"]+)"')
            if ($objects.Count -eq 0) { continue }
            $text = $xml -replace '<w:tab\b[^>]*/>', [string][char]9
            $text = ($text -replace '<[^>]+>', '')
            $text = [System.Net.WebUtility]::HtmlDecode($text)
            $numbered = ($xml -match 'MTPlaceRef')
            foreach ($object in $objects) {
                $ordinal++
                $target = $targets[$object.Groups[1].Value]
                $tex = $null
                $progId = $null
                if (-not [string]::IsNullOrWhiteSpace($target)) {
                    $bytes = Get-ZipEntryBytes -Zip $zip -Name ('word/' + $target)
                    if ($null -ne $bytes) { $tex = Get-MathTypeTexFromOleBytes -Bytes $bytes }
                }
                $records.Add([ordered]@{
                    ordinal        = $ordinal
                    paragraph      = $paragraphIndex
                    ole            = $target
                    numbered       = $numbered
                    layout         = 'unknown'
                    tex            = $tex
                    tex_source     = if ($null -ne $tex) { 'ole-cache' } else { 'missing' }
                    paragraph_text = $text
                })
            }
        }
        return $records
    }
    finally { $zip.Dispose() }
}

function Get-WordToggleSegments {
    param([Parameter(Mandatory)][string]$DocumentPath)
    # Toggle replaces each MathType object with delimited TeX text. Two traps:
    #   1. A delimiter regex over the whole document cannot recover the text reliably,
    #      because a display body may itself contain `\[` (e.g. `\left[`), which ends the
    #      match early and makes the recovered count disagree with the package.
    #   2. Character offsets recorded before a toggle drift, because every toggle changes
    #      the document length, so sampling post-toggle text at a pre-toggle offset lands
    #      inside unrelated prose.
    # Toggling one object at a time and reading the replaced span immediately is the
    # pattern the edit action already relies on, so it is exact by construction.
    $segments = [System.Collections.Generic.List[object]]::new()
    Start-Word -DocumentPath $DocumentPath -LoadMathTypeCommands
    try {
        while ($true) {
            $target = $null
            for ($index = 1; $index -le $script:Document.InlineShapes.Count; $index++) {
                $shape = $script:Document.InlineShapes.Item($index)
                $progId = $null
                try { $progId = $shape.OLEFormat.ProgID } catch {}
                if ($progId -eq 'Equation.DSMT4') { $target = $shape; break }
                Release-ComObject $shape
            }
            if ($null -eq $target) { break }
            $start = [int]$target.Range.Start
            $target.Range.Select()
            $null = $script:Word.Run('MTCommand_TeXToggle')
            Release-ComObject $target
            # The replacement starts at the object's original start position, and the
            # object itself is gone afterwards, so this window is stable.
            $limit = [Math]::Min([int]$script:Document.Content.End, $start + 8000)
            $window = $script:Document.Range($start, $limit).Text
            $match = [regex]::Match($window, '^(\\\[[\s\S]*?\\\]|\$[^$]*\$)')
            if ($match.Success) {
                $rawText = $match.Value
                if ($rawText.StartsWith('\[')) {
                    $segments.Add([ordered]@{ layout = 'display'; tex = $rawText.Substring(2, $rawText.Length - 4) })
                }
                else {
                    $segments.Add([ordered]@{ layout = 'inline'; tex = $rawText.Substring(1, $rawText.Length - 2) })
                }
            }
            else {
                $segments.Add([ordered]@{ layout = 'unknown'; tex = $window.Trim() })
            }
        }
    }
    finally { Stop-Word }
    return $segments
}

function Get-UnsupportedTexCommand {
    <#
    Returns the command names in $Tex that MathType's Toggle TeX cannot convert.

    Detection rule: after a successful conversion, toggling an equation back to TeX must
    never reproduce these names, because the converter replaces every command it
    understands. A name that survives the round trip was echoed verbatim as literal red
    text, or (for \mathop, \lceil, \rceil, \lvert, \rvert) silently produced a wrong or
    missing symbol. Measured against MathType 7.8; do not treat this list as exhaustive.
    #>
    param([string]$Tex)
    if ([string]::IsNullOrWhiteSpace($Tex)) { return @() }
    $known = @(
        '\mathop', '\widetilde', '\Longrightarrow', '\iff', '\lceil', '\rceil',
        '\lvert', '\rvert', '\middle', '\dfrac', '\tfrac', '\text',
        '\begin{cases}', '\begin{align}', '\substack', '\overset', '\underset',
        '\xrightarrow', '\mathbfit', '\bm', '\mathbold', '\widebar',
        '\limits', '\nolimits', '\displaystyle', '\operatorname'
    )
    $found = [System.Collections.Generic.List[string]]::new()
    foreach ($command in $known) {
        if ($Tex.Contains($command)) { $found.Add($command.TrimStart('\')) }
    }
    return @($found)
}

function Get-DocumentBaselineCounts {
    $objects = 0
    for ($index = 1; $index -le $script:Document.InlineShapes.Count; $index++) {
        $shape = $script:Document.InlineShapes.Item($index)
        $progId = $null
        try { $progId = $shape.OLEFormat.ProgID } catch {}
        if ($progId -eq 'Equation.DSMT4') { $objects++ }
        Release-ComObject $shape
    }
    $numbers = 0
    $references = 0
    for ($index = 1; $index -le $script:Document.Fields.Count; $index++) {
        $field = $script:Document.Fields.Item($index)
        $code = $field.Code.Text
        if ($code -match '^\s*MACROBUTTON\s+MTPlaceRef\b') { $numbers++ }
        elseif ($code -match '^\s*GOTOBUTTON\s+(ZEqnNum\d+)\b') { $references++ }
        Release-ComObject $field
    }
    return [ordered]@{
        mathtype_objects        = $objects
        native_number_fields    = $numbers
        native_reference_fields = $references
    }
}

function Invoke-ProseReplacements {
    param($Entries)
    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @($Entries)) {
        $old = [string](Get-PropertyValue -Object $entry -Name 'find' -Default '')
        $new = [string](Get-PropertyValue -Object $entry -Name 'replace' -Default '')
        if ([string]::IsNullOrWhiteSpace($old)) { throw 'Every prose entry requires a non-empty find string.' }
        $before = ([regex]::Matches($script:Document.Content.Text, [regex]::Escape($old))).Count
        $executed = $false
        if ($before -gt 0) {
            $range = $script:Document.Content
            $find = $range.Find
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()
            $executed = [bool]$find.Execute($old, $true, $false, $false, $false, $false, $true, 1, $false, $new, 2)
            Release-ComObject $find
            Release-ComObject $range
        }
        $after = ([regex]::Matches($script:Document.Content.Text, [regex]::Escape($old))).Count
        $expected = Get-PropertyValue -Object $entry -Name 'count'
        $results.Add([ordered]@{
            find      = $old
            replace   = $new
            before    = $before
            after     = $after
            executed  = $executed
            expected  = $expected
            satisfied = (($null -eq $expected) -or ([int]$expected -eq $before))
        })
    }
    return $results
}

# ------------------------------------------------------------------ manifests

function Read-EditManifest {
    param([Parameter(Mandatory)][string]$Path)
    $resolved = Resolve-ExistingFile -Path $Path -Label 'ManifestPath'
    $manifest = Get-Content -Raw -LiteralPath $resolved | ConvertFrom-Json
    if ([int](Get-PropertyValue -Object $manifest -Name 'schema_version' -Default 0) -ne 1) { throw 'Edit manifest schema_version must be 1.' }
    if ($null -eq $manifest.PSObject.Properties['equations']) { throw 'Edit manifest must contain an equations array.' }
    return $manifest
}

function Assert-EditManifest {
    param([Parameter(Mandatory)]$Manifest)
    $indices = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($equation in @($Manifest.equations)) {
        $index = [int](Get-PropertyValue -Object $equation -Name 'index' -Default 0)
        if ($index -lt 1) { throw 'Every edit entry requires a 1-based "index" reported by the inventory action.' }
        if (-not $indices.Add($index)) { throw ('Duplicate edit index: ' + $index) }
        $delete = [bool](Get-PropertyValue -Object $equation -Name 'delete' -Default $false)
        $newTex = [string](Get-PropertyValue -Object $equation -Name 'new_tex' -Default '')
        if ($delete -and -not [string]::IsNullOrWhiteSpace($newTex)) { throw ('Edit ' + $index + ' cannot combine delete with new_tex.') }
        if (-not $delete -and [string]::IsNullOrWhiteSpace($newTex)) { throw ('Edit ' + $index + ' requires new_tex or delete: true.') }
        if ($newTex -match '[\r\n]') { throw ('Edit ' + $index + ' new_tex must stay on one line.') }
    }
}

# ------------------------------------------------------------------ converters

function Find-OfficeConverter {
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($env:MATHTYPE_SOFFICE_PATH)) { $candidates.Add($env:MATHTYPE_SOFFICE_PATH) }
    foreach ($name in @('soffice.exe', 'soffice.com', 'soffice')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command) { $candidates.Add($command.Source) }
    }
    foreach ($root in @($env:ProgramFiles, [Environment]::GetFolderPath('ProgramFilesX86'), 'D:\')) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $candidates.Add((Join-Path $root 'LibreOffice\program\soffice.exe'))
    }
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).Path }
    }
    return $null
}

function Find-PdfToPpm {
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($env:MATHTYPE_PDFTOPPM_PATH)) { $candidates.Add($env:MATHTYPE_PDFTOPPM_PATH) }
    $command = Get-Command pdftoppm.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $command) { $candidates.Add($command.Source) }
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).Path }
    }
    return $null
}

# ------------------------------------------------------------------ install helpers

function Install-MathTypeWordTemplate {
    $template = Find-MathTypeWordTemplate -ExplicitPath $script:TemplatePath
    if ([string]::IsNullOrWhiteSpace($template)) { throw 'MathType Commands (*.dotm) was not found; install MathType 7 for Word first.' }
    $startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
    if (-not (Test-Path -LiteralPath $startup -PathType Container)) { New-Item -ItemType Directory -Path $startup -Force | Out-Null }
    $destination = Join-Path $startup ([IO.Path]::GetFileName($template))
    $backup = $null
    if ((Test-Path -LiteralPath $destination -PathType Leaf) -and ($destination -ne $template)) {
        $backup = $destination + '.backup-' + (Get-Date).ToString('yyyyMMddHHmmss')
        Move-Item -LiteralPath $destination -Destination $backup -Force
    }
    if ($destination -ne $template) { Copy-Item -LiteralPath $template -Destination $destination -Force }
    return [ordered]@{
        source      = $template
        destination = $destination
        backup      = $backup
        startup     = $startup
        already_installed = ($destination -eq $template)
    }
}

# ------------------------------------------------------------------ actions

function Invoke-InstallTemplate {
    $result = Install-MathTypeWordTemplate
    Write-JsonResult ([ordered]@{
        ok          = $true
        action      = 'install-template'
        source      = $result.source
        destination = $result.destination
        backup      = $result.backup
        note        = 'Word loads MathType Commands from its STARTUP folder at every start; restart Word before the next automation run.'
    })
}

function Invoke-Inventory {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'inventory requires InputPath.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($input) -ine '.docx') { throw 'InputPath must be a .docx file.' }
    $records = Get-PackageEquationInventory -Path $input
    $layoutResolved = $false
    $layoutError = $null
    if ($ReadTexFromMathType -and $records.Count -gt 0) {
        $tempDirectory = Join-Path ([IO.Path]::GetTempPath()) ('mathtype-for-word-inventory-' + $script:RunToken)
        New-Item -ItemType Directory -Path $tempDirectory -Force | Out-Null
        try {
            $copy = Join-Path $tempDirectory ([IO.Path]::GetFileName($input))
            Copy-Item -LiteralPath $input -Destination $copy -Force
            $segments = Get-WordToggleSegments -DocumentPath $copy
            if ($segments.Count -ne $records.Count) {
                $layoutError = 'Word reported ' + $segments.Count + ' equations but the package lists ' + $records.Count + '; layout and recovered TeX were skipped.'
            }
            else {
                for ($position = 0; $position -lt $records.Count; $position++) {
                    $records[$position]['layout'] = [string]$segments[$position]['layout']
                    $segmentTex = [string]$segments[$position]['tex']
                    if (-not [string]::IsNullOrWhiteSpace($segmentTex)) {
                        $records[$position]['tex'] = $segmentTex
                        $records[$position]['tex_source'] = 'mathtype-toggle'
                    }
                }
                $layoutResolved = $true
            }
        }
        catch { $layoutError = $_.Exception.Message }
        finally { Remove-Item -LiteralPath $tempDirectory -Recurse -Force -ErrorAction SilentlyContinue }
    }
    if (-not $IncludeText) { foreach ($record in $records) { $record.Remove('paragraph_text') } }
    $numbered = 0
    foreach ($record in $records) { if ([bool]$record['numbered']) { $numbered++ } }
    Write-JsonResult ([ordered]@{
        ok              = $true
        action          = 'inventory'
        document_path   = $input
        equation_count  = $records.Count
        numbered_count  = $numbered
        layout_resolved = $layoutResolved
        layout_error    = $layoutError
        hint            = 'Ordinals are 1-based document order. Pass -ReadTexFromMathType for exact TeX plus inline/display layout, then feed these indices to the edit action.'
        equations       = @($records)
    })
}

function Invoke-Edit {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'edit requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) { throw 'edit requires ManifestPath.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($input) -ine '.docx') { throw 'InputPath must be a .docx file.' }
    $destination = if ([string]::IsNullOrWhiteSpace($OutputPath)) { $input } else { Resolve-OutputFile -Path $OutputPath }
    if (($destination -eq $input) -and (-not $Overwrite)) {
        throw 'Editing in place requires -Overwrite; otherwise pass a distinct -OutputPath so the source stays intact.'
    }
    $manifest = Read-EditManifest -Path $ManifestPath
    Assert-EditManifest -Manifest $manifest
    $edits = @{}
    foreach ($equation in @($manifest.equations)) {
        $edits[[int](Get-PropertyValue -Object $equation -Name 'index')] = $equation
    }
    $report = [ordered]@{
        ok                 = $true
        action             = 'edit'
        input_path         = $input
        output_path        = $destination
        rewritten          = 0
        deleted            = 0
        untouched          = 0
        template_load_mode = $null
        prose              = @()
        problems           = @()
        validation         = $null
    }
    $used = [System.Collections.Generic.HashSet[int]]::new()
    Set-TemporaryMathTypePreferences
    try {
        Start-Word -DocumentPath $input -LoadMathTypeCommands
        $report['template_load_mode'] = [string]$script:TemplateLoadMode
        $baseline = Get-DocumentBaselineCounts
        $ordinal = 0
        $index = 1
        while ($index -le $script:Document.InlineShapes.Count) {
            $shape = $script:Document.InlineShapes.Item($index)
            $progId = $null
            try { $progId = $shape.OLEFormat.ProgID } catch {}
            if ($progId -ne 'Equation.DSMT4') { Release-ComObject $shape; $index++; continue }
            $ordinal++
            if (-not $edits.ContainsKey($ordinal)) {
                $report['untouched'] = [int]$report['untouched'] + 1
                Release-ComObject $shape
                $index++
                continue
            }
            $edit = $edits[$ordinal]
            $startPosition = $shape.Range.Start
            $shape.Range.Select()
            $null = $script:Word.Run('MTCommand_TeXToggle')
            $limit = [Math]::Min([int]$script:Document.Content.End, $startPosition + 4000)
            $window = $script:Document.Range($startPosition, $limit).Text
            $match = [regex]::Match($window, '^(\\\[[\s\S]*?\\\]|\$[^$]*\$)')
            if (-not $match.Success) {
                $report['problems'] += ('Equation ' + $ordinal + ' could not be read back as TeX.')
                Release-ComObject $shape
                $index++
                continue
            }
            $rawText = $match.Value
            $open = if ($rawText.StartsWith('\[')) { '\[' } else { '$' }
            $close = if ($open -eq '\[') { '\]' } else { '$' }
            $body = $rawText.Substring($open.Length, $rawText.Length - $open.Length - $close.Length)
            $expectedOld = [string](Get-PropertyValue -Object $edit -Name 'old_tex' -Default '')
            if ((-not $SkipVerify) -and (-not [string]::IsNullOrWhiteSpace($expectedOld)) -and ($body -ne $expectedOld)) {
                $restore = $script:Document.Range($startPosition, $startPosition + $rawText.Length)
                $restore.Select()
                $null = $script:Word.Run('MTCommand_TeXToggle')
                $report['problems'] += ('Equation ' + $ordinal + ' does not match old_tex. Expected [' + $expectedOld + '] but found [' + $body + '].')
                [void]$used.Add($ordinal)
                Release-ComObject $shape
                $index++
                continue
            }
            $delete = [bool](Get-PropertyValue -Object $edit -Name 'delete' -Default $false)
            if ($delete) {
                $range = $script:Document.Range($startPosition, $startPosition + $rawText.Length)
                $range.Select()
                $script:Selection.Delete()
                $report['deleted'] = [int]$report['deleted'] + 1
                [void]$used.Add($ordinal)
                Release-ComObject $range
                Release-ComObject $shape
            }
            else {
                $newTex = [string](Get-PropertyValue -Object $edit -Name 'new_tex' -Default '')
                $newText = $open + $newTex + $close
                $range = $script:Document.Range($startPosition, $startPosition + $rawText.Length)
                $range.Select()
                $script:Selection.TypeText($newText)
                $newRange = $script:Document.Range($startPosition, $startPosition + $newText.Length)
                $newRange.Select()
                $null = $script:Word.Run('MTCommand_TeXToggle')
                $report['rewritten'] = [int]$report['rewritten'] + 1
                [void]$used.Add($ordinal)
                Release-ComObject $newRange
                Release-ComObject $range
                Release-ComObject $shape
                $index++
            }
        }
        foreach ($key in @($edits.Keys)) {
            if (-not $used.Contains([int]$key)) { $report['problems'] += ('Equation index ' + $key + ' was not found in the document.') }
        }
        if (@($report['problems']).Count -gt 0) {
            throw ('Edit verification failed: ' + (@($report['problems']) -join ' '))
        }
        $proseEntries = @(Get-PropertyValue -Object $manifest -Name 'prose' -Default @())
        if ($proseEntries.Count -gt 0) {
            $proseResults = @(Invoke-ProseReplacements -Entries $proseEntries)
            $report['prose'] = $proseResults
            foreach ($proseResult in $proseResults) {
                if (($null -ne $proseResult['expected']) -and (-not [bool]$proseResult['satisfied'])) {
                    throw ('Prose replacement count mismatch for [' + $proseResult['find'] + ']: expected ' + $proseResult['expected'] + ' but found ' + $proseResult['before'] + '.')
                }
            }
        }
        $script:Document.Fields.Update() | Out-Null
        $validationManifest = [pscustomobject]@{
            baseline_counts = [pscustomobject]@{
                mathtype_objects        = ([int]$baseline['mathtype_objects'] - [int]$report['deleted'])
                native_number_fields    = [int]$baseline['native_number_fields']
                native_reference_fields = [int]$baseline['native_reference_fields']
            }
            equations  = @()
            references = @()
        }
        $report['validation'] = Save-DocumentAtomically -Destination $destination -Manifest $validationManifest
    }
    finally {
        Stop-Word
        Restore-MathTypePreferences
    }
    Write-JsonResult $report
}

function Invoke-Preview {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'preview requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'preview requires OutputPath, the PDF path to write.' }
    $input = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    $pdf = [IO.Path]::GetFullPath($OutputPath)
    if ([IO.Path]::GetExtension($pdf) -ine '.pdf') { throw 'OutputPath must use the .pdf extension.' }
    $directory = Split-Path -Parent $pdf
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
    $converter = Find-OfficeConverter
    if ([string]::IsNullOrWhiteSpace($converter)) { throw 'LibreOffice (soffice) was not found. Install it or set MATHTYPE_SOFFICE_PATH.' }
    $profileRoot = Join-Path ([IO.Path]::GetTempPath()) ('mathtype-for-word-preview-' + $script:RunToken)
    New-Item -ItemType Directory -Path $profileRoot -Force | Out-Null
    $images = @()
    $produced = Join-Path $directory ([IO.Path]::GetFileNameWithoutExtension($input) + '.pdf')
    foreach ($stale in @($produced, $pdf)) { Remove-Item -LiteralPath $stale -Force -ErrorAction SilentlyContinue }
    try {
        $profileUri = 'file:///' + ($profileRoot -replace '\\', '/')
        $convertOutput = & $converter --headless --norestore ('-env:UserInstallation=' + $profileUri) --convert-to pdf --outdir $directory $input 2>&1
        # soffice.exe is a launcher: on Windows it can return before the conversion finishes.
        $deadline = [DateTime]::UtcNow.AddSeconds($script:PreviewTimeoutSeconds)
        while ((-not (Test-Path -LiteralPath $produced -PathType Leaf)) -and ([DateTime]::UtcNow -lt $deadline)) {
            Start-Sleep -Milliseconds 500
        }
        if (-not (Test-Path -LiteralPath $produced -PathType Leaf)) {
            $message = 'LibreOffice did not produce a PDF within ' + $script:PreviewTimeoutSeconds + ' seconds. Output: ' + (@($convertOutput) -join ' ')
            if (-not (Test-Path -LiteralPath $converter -PathType Leaf)) { $message += ' The converter path is no longer valid.' }
            throw $message
        }
        if ($produced -ne $pdf) { Move-Item -LiteralPath $produced -Destination $pdf -Force }
        if (-not [string]::IsNullOrWhiteSpace($ImageDirectory)) {
            $pdfToPpm = Find-PdfToPpm
            if ([string]::IsNullOrWhiteSpace($pdfToPpm)) { throw 'pdftoppm was not found; install Poppler or TeX Live, or drop -ImageDirectory.' }
            if (-not (Test-Path -LiteralPath $ImageDirectory -PathType Container)) { New-Item -ItemType Directory -Path $ImageDirectory -Force | Out-Null }
            $arguments = @('-png', '-r', [string]$Dpi)
            if ($FirstPage -gt 0) { $arguments += @('-f', [string]$FirstPage) }
            if ($LastPage -gt 0) { $arguments += @('-l', [string]$LastPage) }
            $arguments += @($pdf, (Join-Path $ImageDirectory 'page'))
            & $pdfToPpm @arguments | Out-Null
            $images = @(Get-ChildItem -LiteralPath $ImageDirectory -Filter 'page*.png' -File -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object { $_.FullName })
        }
    }
    finally { Remove-Item -LiteralPath $profileRoot -Recurse -Force -ErrorAction SilentlyContinue }
    Write-JsonResult ([ordered]@{
        ok            = $true
        action        = 'preview'
        document_path = $input
        pdf_path      = $pdf
        converter     = $converter
        image_count   = @($images).Count
        images        = @($images)
        note          = 'LibreOffice renders the stored MathType preview; it also prints Word field codes and hidden MathType section markers that Word hides. Judge line breaks, spacing, and glyph style, not the field codes.'
    })
}

