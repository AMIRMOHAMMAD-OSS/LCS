#requires -Version 5.1
<#
Collect research reports and small source/result files from NRP, preserving paths.
Original files are only read. Large data, checkpoints and general archives are
listed in the inventory without reading their contents. Each ZIP stays below
40 MiB. Use -Upload to commit and push ONLY the newly created review directory.
#>
[CmdletBinding()]
param(
    [string]$SourcePath = "$env:USERPROFILE\OneDrive\Desktop\NRP",
    [string]$RepoPath = "$env:USERPROFILE\LCS-NeurALPS-transfer",
    [ValidateRange(1,24)][int]$MaxFileMiB = 12,
    [ValidateRange(1,256)][int]$MaxTotalMiB = 96,
    [switch]$Upload
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$utf8 = New-Object System.Text.UTF8Encoding($false)
$sourceRoot = (Resolve-Path -LiteralPath $SourcePath).Path.TrimEnd([char[]]@('\','/'))
if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) { throw 'SourcePath must be a directory.' }
$repoRoot = & git -C $RepoPath rev-parse --show-toplevel
if ($LASTEXITCODE -ne 0) { throw 'RepoPath must point to your existing LCS checkout.' }
$repoRoot = (Resolve-Path -LiteralPath $repoRoot).Path
$origin = & git -C $repoRoot remote get-url origin
if ($LASTEXITCODE -ne 0 -or $origin -notmatch '(?i)github\.com[/:]AMIRMOHAMMAD-OSS/LCS(?:\.git)?/?$') {
    throw 'This checkout must have AMIRMOHAMMAD-OSS/LCS as origin.'
}
$sourcePrefix = $sourceRoot + [IO.Path]::DirectorySeparatorChar
if ($repoRoot.StartsWith($sourcePrefix, [StringComparison]::OrdinalIgnoreCase) -or $repoRoot -eq $sourceRoot) {
    throw 'Keep the LCS checkout outside the source NRP directory.'
}

function Invoke-LcsGit {
    param([string[]]$GitArguments)
    & git -C $repoRoot @GitArguments
    if ($LASTEXITCODE -ne 0) { throw ('Git command failed: ' + ($GitArguments -join ' ')) }
}

if ($Upload) {
    # This environment change is confined to this PowerShell process.
    $env:GIT_LFS_SKIP_SMUDGE = '1'
    Invoke-LcsGit -GitArguments @('pull','--ff-only')
}

$tag = (Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + [Guid]::NewGuid().ToString('N').Substring(0,8)
$relativeOutput = 'neuralps/reports/NRP_review_' + $tag
$outputDirectory = Join-Path $repoRoot $relativeOutput
[void](New-Item -ItemType Directory -Path $outputDirectory)
Write-Host ('Scanning: ' + $sourceRoot)

$extensions = @('.txt','.md','.pdf','.docx','.doc','.rtf','.json','.jsonl',
    '.csv','.tsv','.xlsx','.xls','.yaml','.yml','.toml','.ini','.cfg','.log',
    '.out','.err','.py','.sh','.ps1','.ipynb','.r','.html','.htm','.js','.css',
    '.tex','.bib','.xml','.png','.jpg','.jpeg','.svg','.pptx')
$excludedDirectories = @('.git','.venv','venv','env','node_modules','__pycache__',
    'site-packages','.cache','.conda')
$inventory = New-Object 'System.Collections.Generic.List[object]'
$candidates = New-Object 'System.Collections.Generic.List[object]'
$scanErrors = New-Object 'System.Collections.Generic.List[object]'
$directories = New-Object 'System.Collections.Generic.Stack[string]'
$directories.Push($sourceRoot)
$maxFileBytes = [long]$MaxFileMiB * 1MB
$maxTotalBytes = [long]$MaxTotalMiB * 1MB

while ($directories.Count -gt 0) {
    $directory = $directories.Pop()
    try { $children = @(Get-ChildItem -LiteralPath $directory -Force) }
    catch {
        $scanErrors.Add([pscustomobject]@{path=$directory; error=$_.Exception.Message})
        continue
    }
    foreach ($item in $children) {
        $relative = $item.FullName.Substring($sourcePrefix.Length).Replace('\','/')
        # Real filesystem links are not followed. OneDrive placeholder files
        # are not excluded merely because they have the ReparsePoint attribute.
        $isLink = $item.LinkType -in @('SymbolicLink','Junction','HardLink')
        if ($item.PSIsContainer) {
            if ($isLink -or $item.Name.StartsWith('.') -or $item.Name -in $excludedDirectories) {
                $inventory.Add([pscustomobject]@{relative_path=$relative; kind='directory'; bytes=$null;
                    modified_utc=$item.LastWriteTimeUtc.ToString('o'); status='DIRECTORY_NOT_TRAVERSED';
                    archive=$null; entry=$null; sha256=$null; duplicate_of=$null; error=$null})
            } else { $directories.Push($item.FullName) }
            continue
        }
        $record = [pscustomobject]@{relative_path=$relative; kind='file'; bytes=[long]$item.Length;
            modified_utc=$item.LastWriteTimeUtc.ToString('o'); status='INVENTORY_ONLY_FILE_TYPE';
            archive=$null; entry=$null; sha256=$null; duplicate_of=$null; error=$null}
        $inventory.Add($record)
        if ($isLink -or $item.Name.StartsWith('.') -or
            $item.Name -match '(?i)^(credentials\.json|client_secret.*\.json|token\.json|access_token\.txt)$') {
            $record.status = 'EXCLUDED_LINK_OR_PRIVATE_CONFIG'
            continue
        }
        $knownArchive = $item.Name -match '(?i)^(NeurALPS_architecture_audit_.*|NeurALPS_Pretraining_Metrics_.*|context_benefit_inspection.*)\.zip$'
        $compressedTable = $item.Name -match '(?i)\.(json|jsonl|csv|tsv)\.gz$'
        $allowed = $item.Extension.ToLowerInvariant() -in $extensions -or $knownArchive -or $compressedTable
        if (-not $allowed) { continue }
        if ($item.Length -gt $maxFileBytes) { $record.status='FILE_SIZE_LIMIT'; continue }
        $priority = 3
        if ($item.Extension.ToLowerInvariant() -in @('.py','.sh','.ps1','.r','.ipynb')) { $priority=2 }
        if ($item.Name -match '(?i)report|summary|protocol|readme|manifest|provenance|validation|integrity|config|split|fold|metric|architect|method|research_question|final_design') { $priority=1 }
        if ($knownArchive -or $item.Name -match '(?i)^(NeurALPS_.*\.(txt|pdf)|integrity\.json|transfer_results\.json|mlp_results\.json|protocols_and_architectures\.json|CATALOG_VALIDATION\.json)$') { $priority=0 }
        $record.status = 'PENDING'
        $candidates.Add([pscustomobject]@{Priority=$priority; Bytes=[long]$item.Length; Relative=$relative;
            FullName=$item.FullName; ModifiedTicks=$item.LastWriteTimeUtc.Ticks; Record=$record})
    }
}

Write-Host ('Inventoried {0} entries; {1} small report/source candidates.' -f $inventory.Count,$candidates.Count)
$state = [pscustomobject]@{Index=0; Reserved=0L; Zip=$null; Stream=$null; Name=$null;
    Paths=(New-Object 'System.Collections.Generic.List[string]')}

function Close-ReviewZip {
    if ($null -ne $state.Zip) { $state.Zip.Dispose(); $state.Zip=$null }
    if ($null -ne $state.Stream) { $state.Stream.Dispose(); $state.Stream=$null }
}

function Add-ReviewEntry {
    param([string]$EntryName, [byte[]]$Data)
    # Reserve room for incompressible data and ZIP directory/header overhead.
    $cost = [long]$Data.Length + [long][math]::Ceiling($Data.Length / 1000.0) +
        2L * $utf8.GetByteCount($EntryName) + 512L
    if ($cost -gt 32MB) { throw ('Entry unexpectedly exceeds part budget: ' + $EntryName) }
    if ($null -eq $state.Zip -or $state.Reserved + $cost -gt 32MB) {
        Close-ReviewZip
        $state.Index++
        $state.Name = 'reports_part{0:D3}.zip' -f $state.Index
        $partPath = Join-Path $outputDirectory $state.Name
        $state.Stream = [IO.File]::Open($partPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write)
        $state.Zip = New-Object IO.Compression.ZipArchive($state.Stream,[IO.Compression.ZipArchiveMode]::Create,$true)
        $state.Reserved=0L
        $state.Paths.Add($partPath)
    }
    $entry = $state.Zip.CreateEntry($EntryName,[IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try { $stream.Write($Data,0,$Data.Length) } finally { $stream.Dispose() }
    $state.Reserved += $cost
    return $state.Name
}

$seen = @{}
$includedBytes = 0L
$includedFiles = 0
$duplicates = 0
$sha = [Security.Cryptography.SHA256]::Create()
try {
    foreach ($candidate in ($candidates | Sort-Object Priority,Bytes,Relative)) {
        $record = $candidate.Record
        if ($includedBytes + $candidate.Bytes -gt $maxTotalBytes) { $record.status='TOTAL_SIZE_LIMIT'; continue }
        try {
            $before = Get-Item -LiteralPath $candidate.FullName
            if ($before.Length -ne $candidate.Bytes -or $before.LastWriteTimeUtc.Ticks -ne $candidate.ModifiedTicks) {
                $record.status='CHANGED_DURING_COLLECTION'; continue
            }
            [byte[]]$raw = [IO.File]::ReadAllBytes($candidate.FullName)
            $after = Get-Item -LiteralPath $candidate.FullName
            if ($after.Length -ne $candidate.Bytes -or $after.LastWriteTimeUtc.Ticks -ne $candidate.ModifiedTicks -or $raw.Length -ne $candidate.Bytes) {
                $record.status='CHANGED_DURING_COLLECTION'; continue
            }
            $digest = ([BitConverter]::ToString($sha.ComputeHash($raw))).Replace('-','').ToLowerInvariant()
        } catch {
            $record.status='READ_ERROR'; $record.error=$_.Exception.Message
            continue
        }
        $record.sha256=$digest
        if ($seen.ContainsKey($digest)) {
            $first = $seen[$digest]
            $record.status='DUPLICATE_CONTENT'; $record.duplicate_of=$first.relative_path
            $record.archive=$first.archive; $record.entry=$first.entry
            $duplicates++
            continue
        }
        $record.entry='files/' + $candidate.Relative
        $record.archive=Add-ReviewEntry -EntryName $record.entry -Data $raw
        $record.status='INCLUDED'
        $seen[$digest]=$record
        $includedBytes += $raw.Length
        $includedFiles++
    }

    # Inventory is included even for excluded weights, raw datasets and archives.
    # Chunk it to keep metadata entries bounded too.
    $inventoryParts = New-Object 'System.Collections.Generic.List[object]'
    for ($offset=0; $offset -lt $inventory.Count; $offset+=1000) {
        $count=[math]::Min(1000,$inventory.Count-$offset)
        $csv=@($inventory.GetRange($offset,$count) | ConvertTo-Csv -NoTypeInformation)
        $name='inventory/files_{0:D4}.csv' -f (1+[int]($offset/1000))
        $data=$utf8.GetBytes(($csv -join "`r`n") + "`r`n")
        $part=Add-ReviewEntry -EntryName $name -Data $data
        $inventoryParts.Add([pscustomobject]@{archive=$part; entry=$name; rows=$count})
    }
} finally {
    Close-ReviewZip
    $sha.Dispose()
}

$archives = @($state.Paths | ForEach-Object {
    $part=Get-Item -LiteralPath $_
    if ($part.Length -ge 40MB) { throw ('Archive unexpectedly exceeds 40 MiB: ' + $part.Name) }
    [pscustomobject]@{name=$part.Name; bytes=$part.Length; sha256=(Get-FileHash -LiteralPath $part.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}
})
$counts=@($inventory | Group-Object status | ForEach-Object { [pscustomobject]@{status=$_.Name; count=$_.Count} })
$manifest=[ordered]@{format='neuralps_windows_review_v1'; created_utc=[DateTime]::UtcNow.ToString('o');
    source_root=$sourceRoot; repository_directory=$relativeOutput; included_files=$includedFiles;
    included_source_bytes=$includedBytes; duplicate_files=$duplicates; inventory_entries=$inventory.Count;
    max_file_mib=$MaxFileMiB; max_total_source_mib=$MaxTotalMiB; archive_limit_mib=40;
    status_counts=$counts; scan_errors=@($scanErrors.ToArray()); archives=$archives;
    inventory_parts=@($inventoryParts.ToArray());
    notes=@('Source files were not modified.','General archives, checkpoints and raw sequence files are inventoried only.',
        'Eligible files were prioritized; omitted files can be requested by their relative paths.',
        'Duplicate content is stored once, with all original paths retained in the inventory.',
        'No project code was executed and no models were loaded.')}
[IO.File]::WriteAllText((Join-Path $outputDirectory 'collection.json'),($manifest | ConvertTo-Json -Depth 12),$utf8)
Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $outputDirectory 'collector.ps1')
Write-Host ('Collected {0} unique files, {1} duplicates, {2:N1} MiB source bytes in {3} ZIP(s).' -f $includedFiles,$duplicates,($includedBytes/1MB),$archives.Count)
Write-Host ('Output: ' + $outputDirectory)
$counts | Format-Table -AutoSize | Out-Host
if ($scanErrors.Count -gt 0) { Write-Warning 'Some directories could not be scanned. See collection.json.' }
if ($includedFiles -eq 0) { throw 'No eligible files could be read. Nothing was uploaded; inspect collection.json.' }

if ($Upload) {
    Invoke-LcsGit -GitArguments @('add','--',$relativeOutput)
    # Commit ONLY this directory, leaving unrelated staged/untracked files alone.
    Invoke-LcsGit -GitArguments @('commit','-m','Add bounded NeurALPS Windows review files','--',$relativeOutput)
    Invoke-LcsGit -GitArguments @('push')
    Write-Host ('Uploaded: ' + $relativeOutput)
    Invoke-LcsGit -GitArguments @('rev-parse','HEAD')
} else {
    Write-Host 'Collection only: no Git commit or push was performed.'
}
