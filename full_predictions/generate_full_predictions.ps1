<#
.SYNOPSIS
    Deterministic ITU BR IFIC (.mdb) Space & Earth Station Interference Screening & Excel (.xlsx) Extractor
    Since AI Hackathon 2026 â€” Traficom Challenge ("Saving Earth from Space Interferers Using Radio Frequencies")

.DESCRIPTION
    Screens ITU BR IFIC SRS Microsoft Access (.mdb) databases (e.g., ific3074.mdb, ific3076.mdb, or any other
    BR IFIC circular database) against Finland's protected national radio stations and Taajuusjakotaulukko
    frequency allocations.

    Automatically re-launches in 32-bit Windows PowerShell (SysWOW64) if invoked from a 64-bit session so that
    the native Windows Microsoft.Jet.OLEDB.4.0 driver can read .mdb files without requiring external dependencies.

.PARAMETER ProjectRoot
    Root directory of the Hackathon2026 project (defaults to parent folder of full_predictions\ or current workspace).

.PARAMETER StationsPath
    Optional path to a protected frequencies dataset (.xlsx, .csv, or .json).
    Accepts raw Excel files (with Finnish or English headers), CSVs with comma/semicolon delimiters,
    or pre-normalized JSON files. No JSON formatting is required from judges!

.PARAMETER OutputDir
    Directory where multi-sheet .xlsx workbooks and .csv extracts are written (defaults to full_predictions\).

.PARAMETER ExtraMdbPaths
    Optional array of additional ITU BR IFIC .mdb file paths to screen for reproducibility with new datasets.

.EXAMPLE
    .\generate_full_predictions.ps1
    # Screens actual_IFIC_data\ific3074\ific3074.mdb and actual_IFIC_data\ific3076\ific3076.mdb and writes .xlsx files.

.EXAMPLE
    .\generate_full_predictions.ps1 -StationsPath "C:\path\to\Any_Protected_Frequencies.xlsx"
    # Directly screens any new raw Excel or CSV frequency table without needing any JSON formatting!

.EXAMPLE
    .\generate_full_predictions.ps1 -ExtraMdbPaths "C:\data\ific3078.mdb"
    # Screens a new ITU BR IFIC .mdb database using the exact same deterministic physics & regulatory engine.
#>
param(
    [string]$ProjectRoot = "",
    [string]$OutputDir = "",
    [string]$StationsPath = "",
    [Alias("MdbPath", "MdbPaths")]
    [string[]]$ExtraMdbPaths = @(),
    [ValidateSet("AUTO", "M4", "M15", "ITU21")]
    [Alias("ScreeningStandard")]
    [string]$Standard = "AUTO",
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# Auto-relaunch in 32-bit PowerShell (SysWOW64) if currently running in 64-bit process (required for Jet.OLEDB.4.0)
if ([Environment]::Is64BitProcess) {
    $wow64Ps = Join-Path $env:windir "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $wow64Ps) {
        Write-Output "Detected 64-bit PowerShell session -> automatically delegating to 32-bit PowerShell (SysWOW64) for Microsoft.Jet.OLEDB.4.0 .mdb access..."
        $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "$PSCommandPath")
        if ($ProjectRoot)  { $argList += @("-ProjectRoot", "$ProjectRoot") }
        if ($OutputDir)    { $argList += @("-OutputDir", "$OutputDir") }
        if ($StationsPath) { $argList += @("-StationsPath", "$StationsPath") }
        if ($Standard)     { $argList += @("-Standard", "$Standard") }
        foreach ($m in $ExtraMdbPaths) { $argList += @("-ExtraMdbPaths", "$m") }
        if ($Force)        { $argList += "-Force" }
        & $wow64Ps @argList
        exit $LASTEXITCODE
    }
}

$scriptDir = Split-Path -Parent $PSCommandPath
$parentDir = Split-Path -Parent $scriptDir

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    # Use script's own folder as root if it contains project assets, or its parent folder if running inside a subfolder (e.g. full_predictions / new_data_trial)
    if ((Test-Path (Join-Path $scriptDir "problem_to_solve")) -or
        (Test-Path (Join-Path $scriptDir "actual_IFIC_data")) -or
        (Test-Path (Join-Path $scriptDir "taajuusjakotaulukko_compiled.js")) -or
        (Test-Path (Join-Path $scriptDir "default_normalized_stations.json"))) {
        $ProjectRoot = $scriptDir
    } elseif ($parentDir -and (
        (Test-Path (Join-Path $parentDir "problem_to_solve")) -or
        (Test-Path (Join-Path $parentDir "actual_IFIC_data")) -or
        (Test-Path (Join-Path $parentDir "taajuusjakotaulukko.json")))) {
        $ProjectRoot = $parentDir
    } else {
        $ProjectRoot = $scriptDir
    }
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    # Default output directory to the script's own folder (or full_predictions if running from the top-level root)
    if ($scriptDir -eq $ProjectRoot -and (Test-Path (Join-Path $ProjectRoot "full_predictions"))) {
        $OutputDir = Join-Path $ProjectRoot "full_predictions"
    } else {
        $OutputDir = $scriptDir
    }
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$totalPipelineStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$pipelineStartTime = Get-Date
Write-Output "Script Start Time: $($pipelineStartTime.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture))"

$outDir = $OutputDir
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

# Determine which MDB datasets to analyze:
# If -ExtraMdbPaths (or -MdbPath) is provided, ONLY run the analysis on those datasets (skip default 3074 & 3076).
$targetMdbJobs = New-Object System.Collections.Generic.List[PSCustomObject]
$runDefaultPair = $false

if ($null -ne $ExtraMdbPaths -and $ExtraMdbPaths.Count -gt 0) {
    Write-Output "Custom MDB dataset path(s) specified -> skipping default BR IFIC 3074 and 3076 datasets."
    foreach ($extraMdb in $ExtraMdbPaths) {
        if ([string]::IsNullOrWhiteSpace($extraMdb)) { continue }
        if (-not (Test-Path $extraMdb)) {
            Write-Output "ERROR: Specified MDB file does not exist: $extraMdb"
            exit 1
        }
        $baseCode = [System.IO.Path]::GetFileNameWithoutExtension($extraMdb).ToUpper()
        $ificCode = if ($baseCode -match '^IFIC[_-]?(\d+)$') { "BR_IFIC_$($Matches[1])" } elseif ($baseCode.StartsWith("BR_")) { $baseCode } else { "BR_$baseCode" }
        $altCode  = "BR_$baseCode"
        $targetMdbJobs.Add([PSCustomObject]@{
            MdbPath   = (Resolve-Path $extraMdb).Path
            IficCode  = $ificCode
            AltCode   = $altCode
            IficTitle = "ITU $ificCode ($extraMdb)"
        })
    }
} else {
    $runDefaultPair = $true
    $mdb3074 = Join-Path $ProjectRoot "actual_IFIC_data\ific3074\ific3074.mdb"
    $mdb3076 = Join-Path $ProjectRoot "actual_IFIC_data\ific3076\ific3076.mdb"
    if (Test-Path $mdb3074) {
        $targetMdbJobs.Add([PSCustomObject]@{
            MdbPath   = $mdb3074
            IficCode  = "BR_IFIC_3074"
            AltCode   = "BR_IFIC3074"
            IficTitle = "ITU BR IFIC 3074 (Space & Earth Station Database)"
        })
    }
    if (Test-Path $mdb3076) {
        $targetMdbJobs.Add([PSCustomObject]@{
            MdbPath   = $mdb3076
            IficCode  = "BR_IFIC_3076"
            AltCode   = "BR_IFIC3076"
            IficTitle = "ITU BR IFIC 3076 (Space & Earth Station Database)"
        })
    }
}

# Pre-flight check: verify if output files for any target dataset already exist in $outDir
if (-not $Force) {
    $existingFilesFound = New-Object System.Collections.Generic.List[string]
    foreach ($job in $targetMdbJobs) {
        $codesToCheck = @($job.IficCode, $job.AltCode) | Select-Object -Unique
        foreach ($c in $codesToCheck) {
            $expectedFiles = @(
                (Join-Path $outDir "$($c)_Full_Predictions.xlsx"),
                (Join-Path $outDir "$($c)_Notice_Executive_Summary.csv"),
                (Join-Path $outDir "$($c)_Finnish_Stations_874_Impact.csv"),
                (Join-Path $outDir "$($c)_Carrier_Station_Link_Hits.csv")
            )
            foreach ($ef in $expectedFiles) {
                if ((Test-Path $ef) -and (-not $existingFilesFound.Contains($ef))) {
                    $existingFilesFound.Add($ef)
                }
            }
        }
    }
    if ($existingFilesFound.Count -gt 0) {
        $totalPipelineStopwatch.Stop()
        $pipelineEndTime = Get-Date
        Write-Output ""
        Write-Output "=========================================================="
        Write-Output "STOPPING: ANALYSIS ALREADY COMPLETED"
        Write-Output "Output files for the requested dataset(s) already exist in:"
        Write-Output "  $outDir"
        Write-Output "Existing output files detected:"
        foreach ($ef in $existingFilesFound) {
            Write-Output "  - $(Split-Path -Leaf $ef)"
        }
        Write-Output ""
        Write-Output "To re-run and overwrite existing files, add -Force or specify a different -OutputDir."
        Write-Output "Start Time:     $($pipelineStartTime.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture))"
        Write-Output "Stop Time:      $($pipelineEndTime.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture))"
        Write-Output "Total Runtime:  $([Math]::Round($totalPipelineStopwatch.Elapsed.TotalSeconds, 2)) seconds"
        Write-Output "=========================================================="
        exit 0
    }
}

Write-Output "1. Loading stations dataset and Taajuusjakotaulukko (Root: $ProjectRoot | Script Dir: $scriptDir) ..."

$searchRoots = @($scriptDir, $ProjectRoot, $parentDir) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path $_) } | Select-Object -Unique

# Auto-discover stations dataset if -StationsPath was not explicitly specified
if ([string]::IsNullOrWhiteSpace($StationsPath)) {
    foreach ($sr in $searchRoots) {
        $candidates = @(
            (Join-Path $sr "Traficom_SinceAI_Frequencies_of_interest_v05112026.xlsx"),
            (Join-Path $sr "problem_to_solve\Traficom_SinceAI_Frequencies_of_interest_v05112026.xlsx"),
            (Join-Path $sr "default_normalized_stations.json"),
            (Join-Path $sr "problem_to_solve\default_normalized_stations.json")
        )
        foreach ($cand in $candidates) {
            if (Test-Path $cand) {
                $StationsPath = $cand
                break
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($StationsPath)) { break }
        $probDir = Join-Path $sr "problem_to_solve"
        if (Test-Path $probDir) {
            $anyXlsx = Get-ChildItem $probDir -Filter "*.xlsx" | Select-Object -First 1
            if ($anyXlsx) { $StationsPath = $anyXlsx.FullName; break }
        }
    }
}

function Parse-FrequencyToMHz($val) {
    if ($null -eq $val -or $val -eq "") { return 0.0 }
    if ($val -is [double] -or $val -is [int] -or $val -is [long]) { return [double]$val }
    $s = $val.ToString().Trim().Replace(',', '.')
    if ($s -match '^([0-9.]+)\s*([a-zA-Z]+)?$') {
        $num = [double]$matches[1]
        $u = if ($matches[2]) { $matches[2].ToLower() } else { "mhz" }
        switch -Regex ($u) {
            '^(ghz|g)$' { return $num * 1000.0 }
            '^(khz|k)$' { return $num / 1000.0 }
            '^(hz)$'    { return $num / 1000000.0 }
            default     { return $num }
        }
    }
    $d = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return 0.0
}

function Get-PropVal($obj, [string[]]$candidates) {
    foreach ($c in $candidates) {
        $p = $obj.PSObject.Properties[$c]
        if ($p -and $null -ne $p.Value -and $p.Value -ne "") { return $p.Value }
    }
    foreach ($c in $candidates) {
        foreach ($p in $obj.PSObject.Properties) {
            if ($p.Name -like "$c*" -and $null -ne $p.Value -and $p.Value -ne "") { return $p.Value }
        }
    }
    return ""
}

function Load-UniversalStations([string]$filePath) {
    Write-Output "   Loading protected stations from: $filePath ..."
    $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
    $rawList = @()

    if ($ext -eq ".xlsx") {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($filePath)
        $sharedStrings = @()
        $ssEntry = $zip.GetEntry("xl/sharedStrings.xml")
        if ($ssEntry) {
            $reader = New-Object System.IO.StreamReader($ssEntry.Open(), [System.Text.Encoding]::UTF8)
            $xml = [xml]$reader.ReadToEnd()
            $reader.Close()
            foreach ($si in $xml.sst.si) {
                $t = if ($si.t) { $si.t } else { ($si.r | ForEach-Object { $_.t }) -join "" }
                $sharedStrings += $t
            }
        }
        $sheetEntry = $zip.GetEntry("xl/worksheets/sheet1.xml")
        $reader = New-Object System.IO.StreamReader($sheetEntry.Open(), [System.Text.Encoding]::UTF8)
        $sheetXml = [xml]$reader.ReadToEnd()
        $reader.Close()
        $zip.Dispose()

        $allRows = @()
        foreach ($row in $sheetXml.worksheet.sheetData.row) {
            $rIdx = [int]$row.r
            $cells = @{}
            foreach ($c in $row.c) {
                $colLetters = $c.r -replace '[0-9]', ''
                $val = if ($c.v) { $c.v } else { "" }
                if ($c.t -eq "s" -and $val -ne "") {
                    $idx = [int]$val
                    if ($idx -lt $sharedStrings.Count) { $val = $sharedStrings[$idx] }
                }
                $cells[$colLetters] = $val
            }
            $allRows += [PSCustomObject]@{ RowNum = $rIdx; Cells = $cells }
        }

        # Auto-detect header row
        $headerRowIdx = -1
        for ($i = 0; $i -lt [Math]::Min(20, $allRows.Count); $i++) {
            $rowVals = $allRows[$i].Cells.Values | ForEach-Object { $_.ToString().ToLower() }
            if (($rowVals -match 'asemanro|station.*number') -and ($rowVals -match 'taajuus|freq')) {
                $headerRowIdx = $i
                break
            }
        }
        if ($headerRowIdx -eq -1) { $headerRowIdx = 0 }

        $headerMap = $allRows[$headerRowIdx].Cells
        Write-Output "   Auto-detected Excel header row at line $($allRows[$headerRowIdx].RowNum)"

        for ($r = $headerRowIdx + 1; $r -lt $allRows.Count; $r++) {
            $cData = $allRows[$r].Cells
            $obj = @{}
            foreach ($col in $headerMap.Keys) {
                $hName = $headerMap[$col].Trim()
                if ($hName) {
                    $obj[$hName] = if ($cData.ContainsKey($col)) { $cData[$col] } else { "" }
                }
            }
            $rawList += (New-Object PSCustomObject -Property $obj)
        }
    } elseif ($ext -eq ".csv" -or $ext -eq ".txt") {
        $rawLines = Get-Content $filePath -Encoding UTF8
        $firstLine = $rawLines[0]
        $delim = if ($firstLine.Contains(";") -and -not $firstLine.Contains(",")) { ";" } else { "," }
        $rawList = Import-Csv $filePath -Delimiter $delim
    } elseif ($ext -eq ".json") {
        $rawList = Get-Content $filePath -Raw -Encoding UTF8 | ConvertFrom-Json
    } else {
        throw "Unsupported stations file format: $filePath (Please provide .xlsx, .csv, or .json)"
    }

    # Normalize records into standardized station schema
    $cleanFloat = {
        param($v)
        if ($null -eq $v -or $v -eq "") { return 0.0 }
        $str = $v.ToString().Trim().Replace(',', '.')
        $res = 0.0
        if ([double]::TryParse($str, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$res)) {
            return $res
        }
        return 0.0
    }

    $normalized = @()
    foreach ($r in $rawList) {
        $stnId   = Get-PropVal $r @("stationId", "Station number", "Asemanro")
        $licence = Get-PropVal $r @("licenceClass", "Class of licence", "Lupalaji")
        $lat     = Get-PropVal $r @("Latitude", "Latitudi", "lat")
        $lon     = Get-PropVal $r @("Longitude", "Longitudi", "lon")
        $gAlt    = Get-PropVal $r @("Ground height (m)", "Maankork (m)", "groundHeightM")
        $mAlt    = Get-PropVal $r @("Mast height (m)", "Mastonkork (m)", "mastHeightM")
        $rxVal   = Get-PropVal $r @("RX frequency", "RX-taajuus", "rxFreqMHz")
        $txVal   = Get-PropVal $r @("TX frequency", "TX-taajuus", "txFreqMHz")
        $bwVal   = Get-PropVal $r @("Channel bandwidth", "Kan.lev.", "bwMHz", "Transmission bandwidth", "LÃ¤hetteen leveys")
        $gain    = Get-PropVal $r @("Ant gain (dB)", "Ant.vahv (dB)", "antGainDbi")
        $azim    = Get-PropVal $r @("Radiation direction", "Max.sÃ¤t.suunta", "Suunta", "azimuthDeg")
        $pol     = Get-PropVal $r @("Polarisation", "Polarisaatio", "pol")

        # Skip sub-header rows (e.g. Row 12 in dual-header Excel files)
        if ($lat.ToString().ToLower() -match 'lat' -or $stnId.ToString().ToLower() -match 'asema|station') {
            continue
        }

        $parsedRx = Parse-FrequencyToMHz $rxVal
        $parsedTx = Parse-FrequencyToMHz $txVal
        $parsedBw = Parse-FrequencyToMHz $bwVal
        if ($parsedBw -le 0) { $parsedBw = 0.025 }

        $normalized += [PSCustomObject]@{
            stationId     = $stnId.ToString().Trim()
            licenceClass  = if ($licence) { $licence.ToString().Trim() } else { "FIXED_LINK" }
            lat           = &$cleanFloat $lat
            lon           = &$cleanFloat $lon
            groundHeightM = &$cleanFloat $gAlt
            mastHeightM   = &$cleanFloat $mAlt
            rxFreqMHz     = $parsedRx
            txFreqMHz     = $parsedTx
            bwMHz         = $parsedBw
            antGainDbi    = &$cleanFloat $gain
            azimuthDeg    = if ($azim -ne "") { &$cleanFloat $azim } else { $null }
            pol           = if ($pol) { $pol.ToString().Trim() } else { "N/A" }
        }
    }
    return $normalized
}

$loadedRawStations = Load-UniversalStations $StationsPath
Write-Output "   Loaded $($loadedRawStations.Count) protected station records."

$taajuusFile = ""
foreach ($sr in $searchRoots) {
    foreach ($rel in @("taajuusjakotaulukko_compiled.js", "problem_to_solve\taajuusjakotaulukko_compiled.js")) {
        $cand = Join-Path $sr $rel
        if (Test-Path $cand) { $taajuusFile = $cand; break }
    }
    if ($taajuusFile) { break }
}
if (-not $taajuusFile) {
    throw "Could not locate taajuusjakotaulukko_compiled.js in $scriptDir or $ProjectRoot"
}
$taajuusRaw = Get-Content $taajuusFile -Raw -Encoding UTF8
$jsonStart = $taajuusRaw.IndexOf('[')
$jsonEnd   = $taajuusRaw.LastIndexOf(']')
$taajuusBands = $taajuusRaw.Substring($jsonStart, $jsonEnd - $jsonStart + 1) | ConvertFrom-Json

function Get-BandInfo([double]$freqMHz) {
    if ($freqMHz -le 0) {
        return [PSCustomObject]@{
            bandStr = "Missing Frequency (0 MHz)"
            services = "Incomplete Record"
            usage = "Blank TX & RX in source Excel"
            owner = "Data Quality Quarantine: Blank Frequency in Source Excel"
            isWifiBt = $false
        }
    }
    $matches = @($taajuusBands | Where-Object { $freqMHz -ge [double]$_.lowMHz -and $freqMHz -le [double]$_.highMHz })
    if ($matches.Count -gt 0) {
        $sorted = $matches | Sort-Object { ([double]$_.highMHz - [double]$_.lowMHz) }
        $best = $sorted[0]
        $wifi = $matches | Where-Object { $_.isWifiBt } | Select-Object -First 1
        return [PSCustomObject]@{
            bandStr = [string]$best.bandStr
            services = [string]$best.services
            usage = [string]$best.usage
            owner = if ($wifi) { [string]$wifi.owner } else { [string]$best.owner }
            isWifiBt = [bool]($null -ne $wifi)
        }
    }
    $isIsm = ($freqMHz -ge 2400 -and $freqMHz -le 2500) -or ($freqMHz -ge 5150 -and $freqMHz -le 5875) -or ($freqMHz -ge 5945 -and $freqMHz -le 6425)
    return [PSCustomObject]@{
        bandStr = "$([Math]::Round($freqMHz-5,1)) - $([Math]::Round($freqMHz+5,1)) MHz"
        services = if ($isIsm) { "MOBILE, FIXED (RLAN/ISM)" } else { "FIXED, MOBILE" }
        usage = if ($isIsm) { "License-exempt RLAN/BT" } else { "National Radiocommunication Service" }
        owner = if ($isIsm) { "License-Exempt: Wi-Fi / Bluetooth (Traficom Reg. 15)" } else { "Traficom Licensed Spectrum Holder" }
        isWifiBt = $isIsm
    }
}

$stations = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $loadedRawStations.Count; $i++) {
    $ns = $loadedRawStations[$i]
    $rx = [double]$ns.rxFreqMHz
    $tx = [double]$ns.txFreqMHz
    $bw = [double]$ns.bwMHz
    if ($bw -le 0) { $bw = 0.025 }
    $primaryFreq = if ($rx -gt 0) { $rx } else { $tx }
    $bInfo = Get-BandInfo $primaryFreq

    $stations.Add([PSCustomObject]@{
        idx = $i
        stationId = [string]$ns.stationId
        licenceClass = if ([string]::IsNullOrWhiteSpace($ns.licenceClass)) { "FIXED_LINK" } else { [string]$ns.licenceClass }
        lat = [double]$ns.lat
        lon = [double]$ns.lon
        groundHeightM = [double]$ns.groundHeightM
        mastHeightM = [double]$ns.mastHeightM
        rxFreqMHz = $rx
        txFreqMHz = $tx
        bwMHz = $bw
        antGainDbi = [double]$ns.antGainDbi
        pol = [string]$ns.pol
        azimuthDeg = $ns.azimuthDeg
        primaryFreqMHz = $primaryFreq
        rxMin = if ($rx -gt 0) { $rx - $bw/2.0 } else { -1.0 }
        rxMax = if ($rx -gt 0) { $rx + $bw/2.0 } else { -1.0 }
        txMin = if ($tx -gt 0) { $tx - $bw/2.0 } else { -1.0 }
        txMax = if ($tx -gt 0) { $tx + $bw/2.0 } else { -1.0 }
        bandInfo = $bInfo
    })
}

# Build unique station frequency intervals for fast pre-filtering
$distinctIntervals = New-Object System.Collections.Generic.List[object]
$seenIv = @{}
foreach ($s in $stations) {
    if ($s.rxMin -gt 0) {
        $k = "$([Math]::Round($s.rxMin,4))_$([Math]::Round($s.rxMax,4))"
        if (-not $seenIv.ContainsKey($k)) {
            $seenIv[$k] = $true
            $distinctIntervals.Add([PSCustomObject]@{ fMin = $s.rxMin; fMax = $s.rxMax })
        }
    }
    if ($s.txMin -gt 0) {
        $k = "$([Math]::Round($s.txMin,4))_$([Math]::Round($s.txMax,4))"
        if (-not $seenIv.ContainsKey($k)) {
            $seenIv[$k] = $true
            $distinctIntervals.Add([PSCustomObject]@{ fMin = $s.txMin; fMax = $s.txMax })
        }
    }
}

# 2. Mathematical Helper Functions (Addressing all Audit Feedback)
function Get-GsoGeometry([double]$rxLatDeg, [double]$rxLonDeg, [double]$satLonDeg) {
    $deg2rad = [Math]::PI / 180.0
    $rad2deg = 180.0 / [Math]::PI
    $latRad = $rxLatDeg * $deg2rad
    $lonDiffRad = ($satLonDeg - $rxLonDeg) * $deg2rad
    $Re = 6371.0
    $Rs = 42157.0 # 6371 + 35786 km

    $cosG = [Math]::Cos($latRad) * [Math]::Cos($lonDiffRad)
    $sinG = [Math]::Sqrt([Math]::Max(0.0, 1.0 - $cosG * $cosG))
    $elevRad = [Math]::Atan2($cosG - ($Re / $Rs), $sinG)
    $elevDeg = $elevRad * $rad2deg
    $slantKm = [Math]::Sqrt([Math]::Max(0.0, $Re*$Re + $Rs*$Rs - 2.0*$Re*$Rs*$cosG))

    # Fixed clockwise-from-North azimuth formula: atan2(sin(dLon), -sin(lat)*cos(dLon))
    $azRad = [Math]::Atan2([Math]::Sin($lonDiffRad), -[Math]::Sin($latRad) * [Math]::Cos($lonDiffRad))
    $azDeg = $azRad * $rad2deg
    if ($azDeg -lt 0) { $azDeg += 360.0 }

    return [PSCustomObject]@{
        elevationDeg = [Math]::Round($elevDeg, 2)
        azimuthDeg   = [Math]::Round($azDeg, 2)
        slantRangeKm = [Math]::Round($slantKm, 1)
        slantRangeM  = $slantKm * 1000.0
        isVisible    = ($elevDeg -gt 0.0)
        geomNote     = if ($elevDeg -gt 0.0) { "GSO Visible ($($satLonDeg)°E)" } else { "Below Horizon in Finland ($($satLonDeg)°E)" }
    }
}

function Get-NgsoGeometry([double]$rxLatDeg, [double]$minAltKm, [double]$maxIncDeg) {
    $Re = 6371.0
    if ($minAltKm -le 0) { $minAltKm = 550.0 }
    if ($maxIncDeg -le 0) { $maxIncDeg = 53.0 }
    $Rs = $Re + $minAltKm
    $maxSubSatLat = if ($maxIncDeg -gt 90.0) { 180.0 - $maxIncDeg } else { $maxIncDeg }

    $deg2rad = [Math]::PI / 180.0
    $rad2deg = 180.0 / [Math]::PI

    if ($maxSubSatLat -ge $rxLatDeg) {
        # Satellite ground track reaches or exceeds station latitude -> can pass directly overhead (delta=90 deg, R=minAltKm)
        return [PSCustomObject]@{
            elevationDeg = 90.0
            azimuthDeg   = 180.0
            slantRangeKm = [Math]::Round($minAltKm, 1)
            slantRangeM  = $minAltKm * 1000.0
            isVisible    = $true
            geomNote     = "NGSO Overhead Pass Capable (h=$minAltKm km, inc=$maxIncDeg°)"
        }
    } else {
        # Closest sub-satellite point is due south on the station's meridian at latitude maxSubSatLat
        $gammaRad = ($rxLatDeg - $maxSubSatLat) * $deg2rad
        $cosG = [Math]::Cos($gammaRad)
        $sinG = [Math]::Sin($gammaRad)
        $elevRad = [Math]::Atan2($cosG - ($Re / $Rs), $sinG)
        $elevDeg = $elevRad * $rad2deg
        $slantKm = [Math]::Sqrt([Math]::Max(0.0, $Re*$Re + $Rs*$Rs - 2.0*$Re*$Rs*$cosG))
        return [PSCustomObject]@{
            elevationDeg = [Math]::Round($elevDeg, 2)
            azimuthDeg   = 180.0
            slantRangeKm = [Math]::Round($slantKm, 1)
            slantRangeM  = $slantKm * 1000.0
            isVisible    = ($elevDeg -gt 0.0)
            geomNote     = if ($elevDeg -gt 0.0) { "NGSO Closest Approach South (h=$minAltKm km, inc=$maxIncDeg°)" } else { "NGSO Below Horizon at Station Latitude (h=$minAltKm km, inc=$maxIncDeg°)" }
        }
    }
}

# ITU RR Article 21 Table 21-4 Dynamic Threshold with Explicit Applicability Gating
function Get-Itu21Threshold([double]$freqMHz, [double]$elevDeg) {
    $delta = [Math]::Max(0.0, [Math]::Min(90.0, $elevDeg))
    $angleOffset = 0.0
    if ($delta -gt 25.0) { $angleOffset = 10.0 }
    elseif ($delta -gt 5.0) { $angleOffset = 0.5 * ($delta - 5.0) }

    # Explicit ITU RR Table 21-4 band applicability check (Addressing Audit Point 1)
    if ($freqMHz -lt 1525.0) {
        return [PSCustomObject]@{
            hasTable21Limit = $false
            limit4kHzDBW    = $null
            bandRule        = "No ITU Table 21-4 PFD limit <1525 MHz (Governed by RR Art. 9 / App. 5 Coordination Trigger; M4 -144 dBW/m²/4kHz national baseline applied)"
        }
    } elseif ($freqMHz -le 2500.0) {
        return [PSCustomObject]@{
            hasTable21Limit = $true
            limit4kHzDBW    = [Math]::Round(-154.0 + $angleOffset, 2)
            bandRule        = "ITU RR Table 21-4 (1525-2500 MHz L/S-Band: -154 to -144 dBW/m²/4kHz)"
        }
    } elseif ($freqMHz -le 2690.0) {
        $off2690 = if ($delta -gt 25.0) { 15.0 } elseif ($delta -gt 5.0) { 0.75 * ($delta - 5.0) } else { 0.0 }
        return [PSCustomObject]@{
            hasTable21Limit = $true
            limit4kHzDBW    = [Math]::Round(-152.0 + $off2690, 2)
            bandRule        = "ITU RR Table 21-4 (2500-2690 MHz S-Band: -152 to -137 dBW/m²/4kHz)"
        }
    } elseif ($freqMHz -le 7750.0) {
        return [PSCustomObject]@{
            hasTable21Limit = $true
            limit4kHzDBW    = [Math]::Round(-152.0 + $angleOffset, 2)
            bandRule        = "ITU RR Table 21-4 (3400-7750 MHz C-Band: -152 to -142 dBW/m²/4kHz)"
        }
    } elseif ($freqMHz -le 11700.0) {
        return [PSCustomObject]@{
            hasTable21Limit = $true
            limit4kHzDBW    = [Math]::Round(-150.0 + $angleOffset, 2)
            bandRule        = "ITU RR Table 21-4 (7750-11700 MHz X/Ku-Band: -150 to -140 dBW/m²/4kHz)"
        }
    } elseif ($freqMHz -le 15400.0) {
        return [PSCustomObject]@{
            hasTable21Limit = $true
            limit4kHzDBW    = [Math]::Round(-148.0 + $angleOffset, 2)
            bandRule        = "ITU RR Table 21-4 (11.7-15.4 GHz Ku-Band: -148 to -138 dBW/m²/4kHz)"
        }
    } else {
        # Above 15.4 GHz (e.g., 17.7-19.7 GHz, 24.65-27.5 GHz): -115 to -105 dBW/m²/1MHz = -139 to -129 dBW/m²/4kHz
        return [PSCustomObject]@{
            hasTable21Limit = $true
            limit4kHzDBW    = [Math]::Round(-139.0 + $angleOffset, 2)
            bandRule        = "ITU RR Table 21-4 (>15.4 GHz Ka-Band: -115 to -105 dBW/m²/1MHz equiv. -139 to -129 dBW/m²/4kHz)"
        }
    }
}

# OpenXML Multi-Sheet .xlsx Builder
function Escape-XmlText([string]$s) {
    if ($null -eq $s) { return "" }
    return [System.Security.SecurityElement]::Escape($s)
}

function Get-ExcelColName([int]$colIdx1Based) {
    [int]$dividend = $colIdx1Based
    $colName = ""
    while ($dividend -gt 0) {
        [int]$modulo = ($dividend - 1) % 26
        $colName = [char]([int](65 + $modulo)) + $colName
        $dividend = [int][Math]::Floor(($dividend - $modulo) / 26)
    }
    return $colName
}

function Write-MultiSheetXlsx([string]$filePath, [array]$sheets) {
    $targetPath = $filePath
    $fs = $null
    try {
        if (Test-Path $targetPath) { Remove-Item $targetPath -Force }
        $fs = New-Object System.IO.FileStream($targetPath, [System.IO.FileMode]::Create)
    } catch [System.IO.IOException] {
        $dir = Split-Path -Parent $targetPath
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
        $ext = [System.IO.Path]::GetExtension($targetPath)
        $targetPath = Join-Path $dir "${baseName}_NEW${ext}"
        Write-Warning "File '$filePath' is currently locked (likely open in Excel). Saved to '$targetPath' instead."
        if (Test-Path $targetPath) { Remove-Item $targetPath -Force }
        $fs = New-Object System.IO.FileStream($targetPath, [System.IO.FileMode]::Create)
    }
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false)

    function Add-ZipTextEntry($archive, [string]$entryPath, [string]$content) {
        $entry = $archive.CreateEntry($entryPath, [System.IO.Compression.CompressionLevel]::Fastest)
        $stream = $entry.Open()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Close()
    }

    # [Content_Types].xml
    $ctOverrides = ""
    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $idx = $i + 1
        $ctOverrides += "<Override PartName=`"/xl/worksheets/sheet$idx.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
    }
    $contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  $ctOverrides
</Types>
"@
    Add-ZipTextEntry $zip "[Content_Types].xml" $contentTypesXml

    # _rels/.rels
    $relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@
    Add-ZipTextEntry $zip "_rels/.rels" $relsXml

    # xl/workbook.xml & xl/_rels/workbook.xml.rels
    $wbSheets = ""
    $wbRels = ""
    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $idx = $i + 1
        $sName = Escape-XmlText $sheets[$i].Name
        $wbSheets += "<sheet name=`"$sName`" sheetId=`"$idx`" r:id=`"rId$idx`"/>"
        $wbRels += "<Relationship Id=`"rId$idx`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$idx.xml`"/>"
    }
    $styleRelId = $sheets.Count + 1
    $wbRels += "<Relationship Id=`"rId$styleRelId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>"

    $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>$wbSheets</sheets>
</workbook>
"@
    Add-ZipTextEntry $zip "xl/workbook.xml" $workbookXml

    $workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  $wbRels
</Relationships>
"@
    Add-ZipTextEntry $zip "xl/_rels/workbook.xml.rels" $workbookRelsXml

    # xl/styles.xml (Style 0 = normal, Style 1 = bold dark header with slate fill)
    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="10"/><color rgb="FF0F172A"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FFCBD5E1"/></top>
      <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>
"@
    Add-ZipTextEntry $zip "xl/styles.xml" $stylesXml

    # Build each worksheet
    for ($sIdx = 0; $sIdx -lt $sheets.Count; $sIdx++) {
        $sheet = $sheets[$sIdx]
        $headers = $sheet.Headers
        $rows = $sheet.Rows
        $colCount = $headers.Count

        # Compute column widths
        $colWidths = New-Object 'double[]' $colCount
        for ($c = 0; $c -lt $colCount; $c++) {
            $colWidths[$c] = [Math]::Max(12, ($headers[$c].ToString().Length + 4))
        }
        $sampleLimit = [Math]::Min($rows.Count, 300)
        for ($r = 0; $r -lt $sampleLimit; $r++) {
            $rowArr = $rows[$r]
            for ($c = 0; $c -lt [Math]::Min($colCount, $rowArr.Count); $c++) {
                $valStr = if ($null -eq $rowArr[$c]) { "" } else { $rowArr[$c].ToString() }
                $len = [Math]::Min(55, $valStr.Length + 3)
                if ($len -gt $colWidths[$c]) { $colWidths[$c] = $len }
            }
        }

        $sb = New-Object System.Text.StringBuilder(65536)
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        [void]$sb.Append('<sheetViews><sheetView tabSelected="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        [void]$sb.Append('<cols>')
        for ($c = 0; $c -lt $colCount; $c++) {
            $c1 = $c + 1
            $w = [Math]::Round($colWidths[$c], 1).ToString([System.Globalization.CultureInfo]::InvariantCulture)
            [void]$sb.Append("<col min=`"$c1`" max=`"$c1`" width=`"$w`" customWidth=`"1`"/>")
        }
        [void]$sb.Append('</cols><sheetData>')

        # Header row (Row 1)
        [void]$sb.Append('<row r="1" ht="20" customHeight="1">')
        for ($c = 0; $c -lt $colCount; $c++) {
            $cellRef = "$(Get-ExcelColName ($c + 1))1"
            $escaped = Escape-XmlText ([string]$headers[$c])
            [void]$sb.Append("<c r=`"$cellRef`" s=`"1`" t=`"inlineStr`"><is><t>$escaped</t></is></c>")
        }
        [void]$sb.Append('</row>')

        # Data rows
        for ($r = 0; $r -lt $rows.Count; $r++) {
            $rowNum = $r + 2
            $rowArr = $rows[$r]
            [void]$sb.Append("<row r=`"$rowNum`">")
            for ($c = 0; $c -lt $colCount; $c++) {
                $val = if ($c -lt $rowArr.Count) { $rowArr[$c] } else { $null }
                if ($null -eq $val -or $val -eq "") { continue }
                $cellRef = "$(Get-ExcelColName ($c + 1))$rowNum"
                if ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal] -or $val -is [float]) {
                    $dVal = [double]$val
                    if (-not [double]::IsNaN($dVal) -and -not [double]::IsInfinity($dVal)) {
                        $numStr = $dVal.ToString([System.Globalization.CultureInfo]::InvariantCulture)
                        [void]$sb.Append("<c r=`"$cellRef`" s=`"0`"><v>$numStr</v></c>")
                        continue
                    }
                }
                $escaped = Escape-XmlText ($val.ToString())
                [void]$sb.Append("<c r=`"$cellRef`" s=`"0`" t=`"inlineStr`"><is><t>$escaped</t></is></c>")
            }
            [void]$sb.Append('</row>')
        }

        $lastColName = Get-ExcelColName $colCount
        $lastRowNum = $rows.Count + 1
        [void]$sb.Append("</sheetData><autoFilter ref=`"A1:$lastColName$lastRowNum`"/></worksheet>")

        $sheetNum = $sIdx + 1
        Add-ZipTextEntry $zip "xl/worksheets/sheet$sheetNum.xml" $sb.ToString()
    }

    $zip.Dispose()
    $fs.Close()
}

function Write-CleanCsv([string]$filePath, [array]$headers, [array]$rows) {
    $sb = New-Object System.Text.StringBuilder(65536)
    $hLine = ($headers | ForEach-Object {
        $s = [string]$_
        if ($s.Contains(",") -or $s.Contains('"')) { '"' + $s.Replace('"', '""') + '"' } else { $s }
    }) -join ","
    [void]$sb.Append("$hLine`r`n")

    foreach ($row in $rows) {
        $cells = New-Object System.Collections.Generic.List[string]
        foreach ($v in $row) {
            if ($null -eq $v) { $cells.Add(""); continue }
            if ($v -is [double] -or $v -is [float] -or $v -is [decimal]) {
                $cells.Add(([double]$v).ToString([System.Globalization.CultureInfo]::InvariantCulture))
            } elseif ($v -is [int] -or $v -is [long]) {
                $cells.Add($v.ToString())
            } else {
                $s = $v.ToString()
                if ($s.Contains(",") -or $s.Contains('"') -or $s.Contains(";") -or $s.Contains("`n") -or $s.Contains("`r")) {
                    $cells.Add('"' + $s.Replace('"', '""') + '"')
                } else {
                    $cells.Add($s)
                }
            }
        }
        [void]$sb.Append(($cells -join ",") + "`r`n")
    }
    try {
        [System.IO.File]::WriteAllText($filePath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    } catch [System.IO.IOException] {
        $dir = Split-Path -Parent $filePath
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
        $ext = [System.IO.Path]::GetExtension($filePath)
        $fallback = Join-Path $dir "${baseName}_NEW${ext}"
        Write-Warning "File '$filePath' is currently locked (likely open in Excel). Saved to '$fallback' instead."
        [System.IO.File]::WriteAllText($fallback, $sb.ToString(), [System.Text.Encoding]::UTF8)
    }
}

Write-Output "Helper functions initialized."


# 3. Core IFIC MDB Screening Function
function Invoke-IficScreening([string]$mdbPath, [string]$ificCode, [string]$ificTitle) {
    $circStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Output "=========================================================="
    Write-Output "PROCESSING $ificTitle ($mdbPath)"
    Write-Output "=========================================================="

    $connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$mdbPath;"
    $conn = New-Object System.Data.OleDb.OleDbConnection($connStr)
    $conn.Open()

    function Get-TableData([string]$sql) {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $sql
        $da = New-Object System.Data.OleDb.OleDbDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        [void]$da.Fill($dt)
        return ,$dt
    }

    # A. Load all notices (union of notice + com_el + geo + non_geo + e_stn)
    $dtNotice  = Get-TableData "SELECT * FROM notice"
    $dtComEl   = Get-TableData "SELECT * FROM com_el"
    $dtGeo     = Get-TableData "SELECT * FROM geo"
    $dtNonGeo  = Get-TableData "SELECT * FROM non_geo"
    $dtEstn    = Get-TableData "SELECT * FROM e_stn"
    $dtPubSsn  = Get-TableData "SELECT * FROM pub_ssn"
    $dtOrbit   = Get-TableData "SELECT * FROM orbit"
    $dtBeam    = Get-TableData "SELECT * FROM s_beam"
    $dtCpfd    = Get-TableData "SELECT * FROM c_pfd"

    # Build lookup maps by ntc_id
    $noticeMap = @{}
    foreach ($r in $dtNotice.Rows) {
        $nid = $r["ntc_id"].ToString()
        $noticeMap[$nid] = [PSCustomObject]@{
            ntc_id      = [long]$r["ntc_id"]
            adm         = $r["adm"].ToString().Trim()
            prov        = $r["prov"].ToString().Trim()
            ntf_rsn     = $r["ntf_rsn"].ToString().Trim()
            act_code    = $r["act_code"].ToString().Trim()
            sat_name    = ""
            ntc_type    = "U"
            long_nom    = $null
            d_rcv       = ""
            ssn_list    = New-Object System.Collections.Generic.List[string]
            min_alt_km  = 550.0
            max_alt_km  = 550.0
            max_inc_deg = 53.0
            has_orbit   = $false
            estn_lat    = $null
            estn_lon    = $null
        }
    }
    foreach ($r in $dtComEl.Rows) {
        $nid = $r["ntc_id"].ToString()
        if (-not $noticeMap.ContainsKey($nid)) {
            $noticeMap[$nid] = [PSCustomObject]@{
                ntc_id      = [long]$r["ntc_id"]
                adm         = $r["adm"].ToString().Trim()
                prov        = $r["prov"].ToString().Trim()
                ntf_rsn     = $r["ntf_rsn"].ToString().Trim()
                act_code    = "A"
                sat_name    = ""
                ntc_type    = "U"
                long_nom    = $null
                d_rcv       = ""
                ssn_list    = New-Object System.Collections.Generic.List[string]
                min_alt_km  = 550.0
                max_alt_km  = 550.0
                max_inc_deg = 53.0
                has_orbit   = $false
                estn_lat    = $null
                estn_lon    = $null
            }
        }
        $nm = $noticeMap[$nid]
        if ($r["sat_name"] -isnot [DBNull]) { $nm.sat_name = $r["sat_name"].ToString().Trim() }
        if ($r["ntc_type"] -isnot [DBNull]) { $nm.ntc_type = $r["ntc_type"].ToString().Trim() }
        if ($r["long_nom"] -isnot [DBNull]) { $nm.long_nom = [Math]::Round([double]$r["long_nom"], 2) }
        if ($r["d_rcv"] -isnot [DBNull])    { $nm.d_rcv = ([datetime]$r["d_rcv"]).ToString("yyyy-MM-dd") }
    }
    foreach ($r in $dtGeo.Rows) {
        $nid = $r["ntc_id"].ToString()
        if ($noticeMap.ContainsKey($nid)) {
            $nm = $noticeMap[$nid]
            $nm.ntc_type = "G"
            if ([string]::IsNullOrWhiteSpace($nm.sat_name) -and $r["sat_name"] -isnot [DBNull]) { $nm.sat_name = $r["sat_name"].ToString().Trim() }
            if ($null -eq $nm.long_nom -and $r["long_nom"] -isnot [DBNull]) { $nm.long_nom = [Math]::Round([double]$r["long_nom"], 2) }
        }
    }
    foreach ($r in $dtNonGeo.Rows) {
        $nid = $r["ntc_id"].ToString()
        if ($noticeMap.ContainsKey($nid)) {
            $nm = $noticeMap[$nid]
            $nm.ntc_type = "N"
            if ([string]::IsNullOrWhiteSpace($nm.sat_name) -and $r["sat_name"] -isnot [DBNull]) { $nm.sat_name = $r["sat_name"].ToString().Trim() }
            $refBody = if ($r["ref_body"] -is [DBNull]) { "T" } else { $r["ref_body"].ToString().Trim() }
            if ($refBody -eq "L") {
                # Lunar orbit (Moon distance ~384,400 km, declination up to 28.5 deg)
                $nm.min_alt_km  = 384400.0
                $nm.max_alt_km  = 384400.0
                $nm.max_inc_deg = 28.5
                $nm.has_orbit   = $true
            } elseif ($refBody -eq "S") {
                # Solar / Lagrange L1/L2 orbit (~1,500,000 km)
                $nm.min_alt_km  = 1500000.0
                $nm.max_alt_km  = 1500000.0
                $nm.max_inc_deg = 23.5
                $nm.has_orbit   = $true
            } elseif ($refBody -eq "J") {
                # Jupiter / Deep Space (~628,730,000 km)
                $nm.min_alt_km  = 628730000.0
                $nm.max_alt_km  = 628730000.0
                $nm.max_inc_deg = 23.5
                $nm.has_orbit   = $true
            }
        }
    }
    foreach ($r in $dtEstn.Rows) {
        $nid = $r["ntc_id"].ToString()
        if ($noticeMap.ContainsKey($nid)) {
            $nm = $noticeMap[$nid]
            $nm.ntc_type = "S"
            if ([string]::IsNullOrWhiteSpace($nm.sat_name) -and $r["stn_name"] -isnot [DBNull]) { $nm.sat_name = $r["stn_name"].ToString().Trim() }
            $latDeg = if ($r["lat_deg"] -is [DBNull]) { 0.0 } else { [double]$r["lat_deg"] }
            $latMin = if ($r["lat_min"] -is [DBNull]) { 0.0 } else { [double]$r["lat_min"] }
            $latSec = if ($r["lat_sec"] -is [DBNull]) { 0.0 } else { [double]$r["lat_sec"] }
            $latSign = if ($r["lat_ns"].ToString().Trim() -eq "S") { -1.0 } else { 1.0 }
            $nm.estn_lat = [Math]::Round($latSign * ($latDeg + $latMin/60.0 + $latSec/3600.0), 4)

            $lonDeg = if ($r["long_deg"] -is [DBNull]) { 0.0 } else { [double]$r["long_deg"] }
            $lonMin = if ($r["long_min"] -is [DBNull]) { 0.0 } else { [double]$r["long_min"] }
            $lonSec = if ($r["long_sec"] -is [DBNull]) { 0.0 } else { [double]$r["long_sec"] }
            $lonSign = if ($r["long_ew"].ToString().Trim() -eq "W") { -1.0 } else { 1.0 }
            $nm.estn_lon = [Math]::Round($lonSign * ($lonDeg + $lonMin/60.0 + $lonSec/3600.0), 4)
        }
    }
    foreach ($r in $dtPubSsn.Rows) {
        $nid = $r["ntc_id"].ToString()
        if ($noticeMap.ContainsKey($nid)) {
            $sRef = "$($r['ssn_ref'].ToString().Trim())/$($r['ssn_no'])"
            if (-not $noticeMap[$nid].ssn_list.Contains($sRef)) {
                $noticeMap[$nid].ssn_list.Add($sRef)
            }
        }
    }
    foreach ($r in $dtOrbit.Rows) {
        $nid = $r["ntc_id"].ToString()
        if ($noticeMap.ContainsKey($nid)) {
            $nm = $noticeMap[$nid]
            $ap = if ($r["apog_km"] -is [DBNull]) { 0.0 } else { [Math]::Round([double]$r["apog_km"], 2) }
            $pe = if ($r["perig_km"] -is [DBNull]) { 0.0 } else { [Math]::Round([double]$r["perig_km"], 2) }
            $op = if ($r["op_ht_km"] -is [DBNull]) { 0.0 } else { [Math]::Round([double]$r["op_ht_km"], 2) }
            $inc = if ($r["inclin_ang"] -is [DBNull]) { 0.0 } else { [Math]::Round([double]$r["inclin_ang"], 2) }
            if ($nm.min_alt_km -ge 380000.0) { continue } # Keep Lunar/Solar/Deep-Space distance if ref_body != 'T'
            $altCandidate = if ($pe -gt 50.0) { $pe } elseif ($op -gt 50.0) { $op } elseif ($ap -gt 50.0) { $ap } else { 550.0 }
            $maxCandidate = if ($ap -gt $altCandidate) { $ap } else { $altCandidate }
            if (-not $nm.has_orbit) {
                $nm.min_alt_km = $altCandidate
                $nm.max_alt_km = $maxCandidate
                $nm.max_inc_deg = if ($inc -gt 0) { $inc } else { 53.0 }
                $nm.has_orbit = $true
            } else {
                if ($altCandidate -lt $nm.min_alt_km) { $nm.min_alt_km = $altCandidate }
                if ($maxCandidate -gt $nm.max_alt_km) { $nm.max_alt_km = $maxCandidate }
                $effIncNew = if ($inc -gt 90.0) { 180.0 - $inc } else { $inc }
                $effIncOld = if ($nm.max_inc_deg -gt 90.0) { 180.0 - $nm.max_inc_deg } else { $nm.max_inc_deg }
                if ($effIncNew -gt $effIncOld) { $nm.max_inc_deg = $inc }
            }
        }
    }

    # Build Beam Gain lookup map: "ntc_id|emi_rcp|beam_name" -> gain dBi
    $beamGainMap = @{}
    foreach ($r in $dtBeam.Rows) {
        $k = "$($r['ntc_id'])|$($r['emi_rcp'].ToString().Trim())|$($r['beam_name'].ToString().Trim())"
        if ($r["gain"] -isnot [DBNull]) {
            $beamGainMap[$k] = [double]$r["gain"]
        }
    }

    # Load aggregated emission power per grp_id from emiss table
    Write-Output "   Aggregating emission power & spectral density from emiss table..."
    $dtEmissAgg = Get-TableData "SELECT grp_id, MAX(pep_max) AS max_pep, MAX(pwr_ds_max) AS max_psd, COUNT(*) AS emi_cnt FROM emiss GROUP BY grp_id"
    $emissMap = @{}
    foreach ($r in $dtEmissAgg.Rows) {
        $gid = $r["grp_id"].ToString()
        $emissMap[$gid] = [PSCustomObject]@{
            max_pep = if ($r["max_pep"] -is [DBNull]) { $null } else { [double]$r["max_pep"] }
            max_psd = if ($r["max_psd"] -is [DBNull]) { $null } else { [double]$r["max_psd"] }
            emi_cnt = [int]$r["emi_cnt"]
        }
    }

    # Load all frequency groups
    Write-Output "   Loading frequency groups from grp table..."
    $dtGrp = Get-TableData "SELECT grp_id, ntc_id, emi_rcp, beam_name, freq_min, freq_max, bdwdth, pwr_max, eirp_nom, srv_code, elev_min, prov FROM grp"
    Write-Output "   Loaded $($dtGrp.Rows.Count) groups across $($noticeMap.Count) notices."

    $conn.Close()

    # Track per-notice counters and deduplicated overlapping downlink carriers
    $noticeStats = @{}
    foreach ($nid in $noticeMap.Keys) {
        $noticeStats[$nid] = [PSCustomObject]@{
            totalDownlinkGrps = 0
            totalUplinkGrps   = 0
            overlapDownGrps   = 0
            overlapUpGrps     = 0
            minFreqMHz        = $null
            maxFreqMHz        = $null
            maxGainDbi        = $null
            maxPepDbw         = $null
            maxPsdDbwHz       = $null
            maxEirp4kHzDbw    = $null
            worstPfd4kHz      = $null
            worstMarginM4     = $null
            worstMarginM15    = $null
            worstMarginItu21  = $null
            overlapStations   = New-Object System.Collections.Generic.HashSet[string]
            exceedStationsM4  = New-Object System.Collections.Generic.HashSet[string]
            exceedStationsItu = New-Object System.Collections.Generic.HashSet[string]
            concernStations   = New-Object System.Collections.Generic.HashSet[string]
            wifiBtStations    = New-Object System.Collections.Generic.HashSet[string]
            licenceClasses    = New-Object System.Collections.Generic.HashSet[string]
            bandOwners        = New-Object System.Collections.Generic.HashSet[string]
            bandStrings       = New-Object System.Collections.Generic.HashSet[string]
            missingPowerGrps  = 0
        }
    }

    $carrierMap = @{} # key: "ntc_id|beam|fMin|fMax" -> aggregated carrier object

    foreach ($r in $dtGrp.Rows) {
        $nid = $r["ntc_id"].ToString()
        if (-not $noticeStats.ContainsKey($nid)) { continue }
        $ns = $noticeStats[$nid]
        $emiRcp = $r["emi_rcp"].ToString().Trim()
        $gMin = [double]$r["freq_min"]
        $gMax = [double]$r["freq_max"]
        $bwMHz = if ($r["bdwdth"] -is [DBNull]) { 0.0 } else { [double]$r["bdwdth"] / 1000.0 }
        if ($gMax -le $gMin -and $bwMHz -gt 0) {
            $c = $gMin; $gMin = $c - $bwMHz/2.0; $gMax = $c + $bwMHz/2.0
        }
        $effBwMHz = if ($bwMHz -gt 0) { $bwMHz } elseif ($gMax -gt $gMin) { $gMax - $gMin } else { 0.025 }

        if ($null -eq $ns.minFreqMHz -or $gMin -lt $ns.minFreqMHz) { $ns.minFreqMHz = $gMin }
        if ($null -eq $ns.maxFreqMHz -or $gMax -gt $ns.maxFreqMHz) { $ns.maxFreqMHz = $gMax }

        if ($emiRcp -eq "E") { $ns.totalDownlinkGrps++ } else { $ns.totalUplinkGrps++ }

        # Check if this group overlaps any Finnish station interval
        $overlapsAny = $false
        foreach ($iv in $distinctIntervals) {
            if ($gMin -le $iv.fMax -and $gMax -ge $iv.fMin) {
                $overlapsAny = $true
                break
            }
        }
        if (-not $overlapsAny) { continue }

        if ($emiRcp -ne "E") {
            $ns.overlapUpGrps++
            continue
        }
        $ns.overlapDownGrps++

        $gid = $r["grp_id"].ToString()
        $beamName = $r["beam_name"].ToString().Trim()
        $beamKey = "$nid|E|$beamName"
        $gainDbi = if ($beamGainMap.ContainsKey($beamKey)) { $beamGainMap[$beamKey] } else { $null }
        $emInfo = if ($emissMap.ContainsKey($gid)) { $emissMap[$gid] } else { $null }
        $pepMax = if ($null -ne $emInfo -and $null -ne $emInfo.max_pep) { $emInfo.max_pep } elseif ($r["pwr_max"] -isnot [DBNull]) { [double]$r["pwr_max"] } else { $null }
        $psdMax = if ($null -ne $emInfo -and $null -ne $emInfo.max_psd) { $emInfo.max_psd } else { $null }

        if ($null -ne $gainDbi -and ($null -eq $ns.maxGainDbi -or $gainDbi -gt $ns.maxGainDbi)) { $ns.maxGainDbi = $gainDbi }
        if ($null -ne $pepMax -and ($null -eq $ns.maxPepDbw -or $pepMax -gt $ns.maxPepDbw)) { $ns.maxPepDbw = $pepMax }
        if ($null -ne $psdMax -and ($null -eq $ns.maxPsdDbwHz -or $psdMax -gt $ns.maxPsdDbwHz)) { $ns.maxPsdDbwHz = $psdMax }

        $ck = "$nid|$beamName|$([Math]::Round($gMin,3))|$([Math]::Round($gMax,3))"
        if (-not $carrierMap.ContainsKey($ck)) {
            $carrierMap[$ck] = [PSCustomObject]@{
                ntc_id    = $nid
                beam_name = $beamName
                freq_min  = $gMin
                freq_max  = $gMax
                bw_mhz    = $effBwMHz
                srv_code  = $r["srv_code"].ToString().Trim()
                prov      = $r["prov"].ToString().Trim()
                gain_dbi  = $gainDbi
                pep_max   = $pepMax
                psd_max   = $psdMax
                grp_count = 1
                sample_grp_id = [long]$r["grp_id"]
            }
        } else {
            $cm = $carrierMap[$ck]
            $cm.grp_count++
            if ($null -ne $gainDbi -and ($null -eq $cm.gain_dbi -or $gainDbi -gt $cm.gain_dbi)) { $cm.gain_dbi = $gainDbi }
            if ($null -ne $pepMax  -and ($null -eq $cm.pep_max  -or $pepMax  -gt $cm.pep_max))  { $cm.pep_max  = $pepMax }
            if ($null -ne $psdMax  -and ($null -eq $cm.psd_max  -or $psdMax  -gt $cm.psd_max))  { $cm.psd_max  = $psdMax }
            if ($effBwMHz -gt $cm.bw_mhz) { $cm.bw_mhz = $effBwMHz }
        }
    }

    Write-Output "   Evaluating $($carrierMap.Count) deduplicated overlapping satellite carriers against $($stations.Count) Finnish stations..."

    # Track per-station impact summary across all carriers in this IFIC circular
    $stationImpacts = @{}
    foreach ($s in $stations) {
        $stationImpacts[$s.idx] = [PSCustomObject]@{
            station            = $s
            overlapNotices     = New-Object System.Collections.Generic.HashSet[string]
            visibleNotices     = New-Object System.Collections.Generic.HashSet[string]
            exceedNoticesM4    = New-Object System.Collections.Generic.HashSet[string]
            exceedNoticesItu21 = New-Object System.Collections.Generic.HashSet[string]
            concernNotices     = New-Object System.Collections.Generic.HashSet[string]
            worstHit           = $null
            allSatNames        = New-Object System.Collections.Generic.HashSet[string]
        }
    }

    $carrierHitsList = New-Object System.Collections.Generic.List[object]

    foreach ($cm in $carrierMap.Values) {
        $nm = $noticeMap[$cm.ntc_id]
        $ns = $noticeStats[$cm.ntc_id]

        # Compute carrier EIRP per 4 kHz and Total Peak EIRP
        # Method 1 (Primary per ITU & context.md): Filed max power spectral density (pwr_ds_max in dBW/Hz) + beam gain + 10*log10(4000 Hz)
        # Method 2 (Fallback): Total PEP (pep_max in dBW) + beam gain - 10*log10(BW_Hz / 4000 Hz)
        $hasPowerData = ($null -ne $cm.gain_dbi) -and (($null -ne $cm.psd_max) -or ($null -ne $cm.pep_max))
        $eirpTotalDbw = $null
        $eirp4kHzPsdDbw = $null
        $eirp4kHzUniDbw = $null
        $eirp4kHzUsedDbw = $null
        $powerMethod = "MISSING_POWER_OR_GAIN"

        if ($hasPowerData) {
            if ($null -ne $cm.pep_max) {
                $eirpTotalDbw = [Math]::Round($cm.pep_max + $cm.gain_dbi, 2)
                $bwKHz = [Math]::Max(4.0, $cm.bw_mhz * 1000.0)
                $eirp4kHzUniDbw = [Math]::Round($eirpTotalDbw - 10.0 * [Math]::Log10($bwKHz / 4.0), 2)
            }
            if ($null -ne $cm.psd_max) {
                # 10*log10(4000 Hz) = 36.0206 dB
                $rawPsd4kHz = $cm.psd_max + $cm.gain_dbi + 36.0206
                if ($null -ne $eirpTotalDbw -and $rawPsd4kHz -gt $eirpTotalDbw) {
                    $rawPsd4kHz = $eirpTotalDbw
                }
                $eirp4kHzPsdDbw = [Math]::Round($rawPsd4kHz, 2)
            }
            if ($null -ne $eirp4kHzPsdDbw) {
                $eirp4kHzUsedDbw = $eirp4kHzPsdDbw
                $powerMethod = "FILED_PSD_MAX (pwr_ds_max + gain)"
            } else {
                $eirp4kHzUsedDbw = $eirp4kHzUniDbw
                $powerMethod = "UNIFORM_PEP_FALLBACK (pep_max + gain)"
            }
            if ($null -eq $ns.maxEirp4kHzDbw -or $eirp4kHzUsedDbw -gt $ns.maxEirp4kHzDbw) {
                $ns.maxEirp4kHzDbw = $eirp4kHzUsedDbw
            }
        } else {
            $ns.missingPowerGrps += $cm.grp_count
        }

        # Screen against all 874 stations
        foreach ($s in $stations) {
            if ($s.primaryFreqMHz -le 0) { continue }

            # Check overlap with RX or TX ("Every TX is received in Finland")
            $ovRx = if ($s.rxMin -gt 0) { [Math]::Min($cm.freq_max, $s.rxMax) - [Math]::Max($cm.freq_min, $s.rxMin) } else { -1.0 }
            $ovTx = if ($s.txMin -gt 0) { [Math]::Min($cm.freq_max, $s.txMax) - [Math]::Max($cm.freq_min, $s.txMin) } else { -1.0 }
            if ($ovRx -lt 0 -and $ovTx -lt 0) { continue }

            $matchedRole = if ($ovRx -ge 0) { "RX" } else { "TX (Traficom Rule: Every TX Received in Finland)" }
            $matchedFreq = if ($ovRx -ge 0) { $s.rxFreqMHz } else { $s.txFreqMHz }
            $overlapMHz  = [Math]::Round([Math]::Max(0.001, [Math]::Max($ovRx, $ovTx)), 4)

            $stImp = $stationImpacts[$s.idx]
            [void]$stImp.overlapNotices.Add($cm.ntc_id)
            [void]$stImp.allSatNames.Add("$($nm.sat_name) [$($nm.adm)]")
            [void]$ns.overlapStations.Add($s.stationId)
            [void]$ns.licenceClasses.Add($s.licenceClass)
            [void]$ns.bandOwners.Add($s.bandInfo.owner)
            [void]$ns.bandStrings.Add($s.bandInfo.bandStr)

            # Calculate 3D Geometry (GSO vs NGSO)
            $geom = if ($nm.ntc_type -eq "G" -and $null -ne $nm.long_nom) {
                Get-GsoGeometry $s.lat $s.lon $nm.long_nom
            } else {
                Get-NgsoGeometry $s.lat $nm.min_alt_km $nm.max_inc_deg
            }

            if ($geom.isVisible) {
                [void]$stImp.visibleNotices.Add($cm.ntc_id)
            }

            # Calculate off-axis angle relative to terrestrial station azimuth if known
            $offAxisDeg = $null
            if ($null -ne $s.azimuthDeg -and $geom.isVisible) {
                $diff = [Math]::Abs($geom.azimuthDeg - $s.azimuthDeg) % 360.0
                if ($diff -gt 180.0) { $diff = 360.0 - $diff }
                $offAxisDeg = [Math]::Round($diff, 1)
            }

            # Calculate Arriving PFD & Regulatory Margins
            $pfd4kHz = $null
            $pfdUni4kHz = $null
            $marginM4 = $null
            $marginM15 = $null
            $itu21Info = Get-Itu21Threshold $matchedFreq $geom.elevationDeg
            $marginItu21 = $null
            $status = "BELOW_HORIZON"

            if ($geom.isVisible) {
                if ($hasPowerData -and $null -ne $eirp4kHzUsedDbw) {
                    $spreadingLossDb = 10.0 * [Math]::Log10(4.0 * [Math]::PI * $geom.slantRangeM * $geom.slantRangeM)
                    $pfd4kHz = [Math]::Round($eirp4kHzUsedDbw - $spreadingLossDb, 2)
                    if ($null -ne $eirp4kHzUniDbw) {
                        $pfdUni4kHz = [Math]::Round($eirp4kHzUniDbw - $spreadingLossDb, 2)
                    }
                    $marginM4  = [Math]::Round($pfd4kHz - (-144.0), 2)
                    $marginM15 = [Math]::Round($pfd4kHz - (-154.0), 2)
                    if ($itu21Info.hasTable21Limit) {
                        $marginItu21 = [Math]::Round($pfd4kHz - $itu21Info.limit4kHzDBW, 2)
                    }

                    if ($null -eq $ns.worstPfd4kHz -or $pfd4kHz -gt $ns.worstPfd4kHz) { $ns.worstPfd4kHz = $pfd4kHz }
                    if ($null -eq $ns.worstMarginM4 -or $marginM4 -gt $ns.worstMarginM4) { $ns.worstMarginM4 = $marginM4 }
                    if ($null -eq $ns.worstMarginM15 -or $marginM15 -gt $ns.worstMarginM15) { $ns.worstMarginM15 = $marginM15 }
                    if ($null -ne $marginItu21 -and ($null -eq $ns.worstMarginItu21 -or $marginItu21 -gt $ns.worstMarginItu21)) { $ns.worstMarginItu21 = $marginItu21 }

                    # Determine status based on selected Standard parameter (AUTO / M4 / M15 / ITU21)
                    $exceedsPrimary = switch ($Standard.ToUpper()) {
                        "M4"    { $marginM4 -gt 0.0 }
                        "M15"   { $marginM15 -gt 0.0 }
                        "ITU21" { ($null -ne $marginItu21 -and $marginItu21 -gt 0.0) }
                        default { ($marginM4 -gt 0.0) -or ($null -ne $marginItu21 -and $marginItu21 -gt 0.0) }
                    }
                    $concernPrimary = switch ($Standard.ToUpper()) {
                        "M4"    { $marginM4 -ge -6.0 }
                        "M15"   { $marginM15 -ge -6.0 }
                        "ITU21" { ($null -ne $marginItu21 -and $marginItu21 -ge -6.0) }
                        default { ($marginM4 -ge -6.0) -or ($null -ne $marginItu21 -and $marginItu21 -ge -6.0) }
                    }

                    if ($s.bandInfo.isWifiBt -and ($exceedsPrimary -or $concernPrimary)) {
                        $status = "LOW_SEVERITY_WIFI_BT"
                        [void]$ns.wifiBtStations.Add($s.stationId)
                    } elseif ($exceedsPrimary) {
                        $status = if ($Standard -eq "M15") { "CRITICAL_M15_EXCEEDED" } elseif (-not $itu21Info.hasTable21Limit) { "CRITICAL_M4_EXCEEDED (NO_ITU21_TABLE_LIMIT_IN_BAND)" } else { "CRITICAL_EXCEEDANCE" }
                        if ($marginM4 -gt 0.0) {
                            [void]$ns.exceedStationsM4.Add($s.stationId)
                            [void]$stImp.exceedNoticesM4.Add($cm.ntc_id)
                        }
                        if ($null -ne $marginItu21 -and $marginItu21 -gt 0.0) {
                            [void]$ns.exceedStationsItu.Add($cm.ntc_id)
                            [void]$stImp.exceedNoticesItu21.Add($cm.ntc_id)
                        }
                    } elseif ($concernPrimary) {
                        $status = "POTENTIAL_CONCERN_WITHIN_6DB"
                        [void]$ns.concernStations.Add($s.stationId)
                        [void]$stImp.concernNotices.Add($cm.ntc_id)
                    } else {
                        $status = if (-not $itu21Info.hasTable21Limit) { "COMPLIANT_M4 (COORD_TRIGGER_BAND)" } else { "COMPLIANT_WITHIN_LIMITS" }
                    }
                } else {
                    $status = "ASSESSMENT_INCOMPLETE_MISSING_POWER"
                }
            }

            $hitObj = [PSCustomObject]@{
                ificCode       = $ificCode
                ntc_id         = [long]$cm.ntc_id
                sat_name       = $nm.sat_name
                adm            = $nm.adm
                ntc_type       = if ($nm.ntc_type -eq "G") { "GSO" } elseif ($nm.ntc_type -eq "N") { "NGSO" } else { $nm.ntc_type }
                orbit_desc     = if ($nm.ntc_type -eq "G") { "GSO $($nm.long_nom)°E" } else { "NGSO h=$($nm.min_alt_km)-$($nm.max_alt_km)km, i=$($nm.max_inc_deg)°" }
                prov           = $nm.prov
                ssn_refs       = ($nm.ssn_list -join "; ")
                beam_name      = $cm.beam_name
                sat_freq_min   = [Math]::Round($cm.freq_min, 3)
                sat_freq_max   = [Math]::Round($cm.freq_max, 3)
                sat_bw_mhz     = [Math]::Round($cm.bw_mhz, 3)
                gain_dbi       = $cm.gain_dbi
                pep_max_dbw    = $cm.pep_max
                psd_max_dbw_hz = $cm.psd_max
                eirp_4khz_dbw  = $eirp4kHzUsedDbw
                power_method   = $powerMethod
                station_id     = $s.stationId
                licence_class  = $s.licenceClass
                lat            = $s.lat
                lon            = $s.lon
                stn_freq_mhz   = $matchedFreq
                stn_bw_mhz     = $s.bwMHz
                matched_role   = $matchedRole
                overlap_mhz    = $overlapMHz
                band_str       = $s.bandInfo.bandStr
                band_owner     = $s.bandInfo.owner
                is_wifi_bt     = if ($s.bandInfo.isWifiBt) { "YES" } else { "NO" }
                elev_deg       = $geom.elevationDeg
                azim_deg       = $geom.azimuthDeg
                stn_azim_deg   = $s.azimuthDeg
                off_axis_deg   = $offAxisDeg
                slant_km       = $geom.slantRangeKm
                pfd_4khz       = $pfd4kHz
                pfd_uni_4khz   = $pfdUni4kHz
                itu21_limit    = $itu21Info.limit4kHzDBW
                itu21_rule     = $itu21Info.bandRule
                margin_itu21   = $marginItu21
                margin_m4      = $marginM4
                margin_m15     = $marginM15
                status         = $status
            }
            $carrierHitsList.Add($hitObj)

            # Update station's worst hit
            if ($null -eq $stImp.worstHit) {
                $stImp.worstHit = $hitObj
            } else {
                $curWorst = $stImp.worstHit
                if ($geom.isVisible -and -not ($curWorst.elev_deg -gt 0)) {
                    $stImp.worstHit = $hitObj
                } elseif ($geom.isVisible -and $null -ne $pfd4kHz) {
                    if ($null -eq $curWorst.pfd_4khz -or $pfd4kHz -gt $curWorst.pfd_4khz) {
                        $stImp.worstHit = $hitObj
                    }
                }
            }
        }
    }

    Write-Output "   Total carrier-to-station link hits evaluated: $($carrierHitsList.Count)"

    # B. Build Sheet 1: 1_Executive_Summary_By_Notice
    $sheet1Headers = @(
        "IFIC_Circular", "Notice_ID", "Satellite_or_Station_Name", "Admin", "Orbit_Type",
        "Provision", "Notification_Reason", "Action_Code", "Special_Section_Refs",
        "Orbital_Slot_or_Altitude", "Horizon_Visibility_in_Finland",
        "Total_Downlink_Groups", "Total_Uplink_Groups", "Overlapping_Downlink_Groups",
        "Freq_Range_MHz", "Overlapping_Finnish_Stations",
        "Max_Beam_Gain_dBi", "Max_PEP_dBW", "Max_PSD_dBW_Hz", "Max_EIRP_4kHz_dBW",
        "Worst_Arriving_PFD_dBW_m2_4kHz", "Worst_Margin_M4_dB (-144)", "Worst_Margin_M15_dB (-154)", "Worst_Margin_ITU_Art21_dB",
        "Stations_Exceeding_M4", "Stations_Exceeding_ITU21", "Stations_Potential_Concern", "Stations_WiFi_BT_Downgraded",
        "Affected_Finnish_Licence_Classes", "Affected_Taajuusjakotaulukko_Owners", "Screening_Disposition"
    )
    $sheet1Rows = New-Object System.Collections.Generic.List[object]
    $noticeSummaryObjects = New-Object System.Collections.Generic.List[object]

    $sortedNotices = $noticeMap.Values | Sort-Object {
        $ns = $noticeStats[$_.ntc_id.ToString()]
        if ($null -ne $ns.worstMarginM4) { -$ns.worstMarginM4 } else { 9999.0 - $ns.overlapDownGrps }
    }

    foreach ($nm in $sortedNotices) {
        $nid = $nm.ntc_id.ToString()
        $ns = $noticeStats[$nid]
        $orbitDesc = if ($nm.ntc_type -eq "G") {
            "GSO $($nm.long_nom)°E"
        } elseif ($nm.ntc_type -eq "N") {
            if ($nm.has_orbit) { "NGSO $($nm.min_alt_km)-$($nm.max_alt_km) km (inc $($nm.max_inc_deg)°)" } else { "NGSO Constellation" }
        } elseif ($nm.ntc_type -eq "S") {
            "Earth Station ($($nm.estn_lat)°N, $($nm.estn_lon)°E)"
        } else {
            "Space Station"
        }

        # Check visibility over Finland (Helsinki 60.17N, 24.94E to Utsjoki 69.90N, 27.02E)
        $visStr = "YES"
        $isVisFin = $true
        if ($nm.ntc_type -eq "G" -and $null -ne $nm.long_nom) {
            $gHel = Get-GsoGeometry 60.17 24.94 $nm.long_nom
            $gUts = Get-GsoGeometry 69.90 27.02 $nm.long_nom
            if ($gHel.isVisible -or $gUts.isVisible) {
                $visStr = "VISIBLE (Elev $($gUts.elevationDeg)° to $($gHel.elevationDeg)°)"
            } else {
                $visStr = "BELOW_HORIZON (Max Elev $($gHel.elevationDeg)° < 0°)"
                $isVisFin = $false
            }
        } elseif ($nm.ntc_type -eq "N") {
            $gNgso = Get-NgsoGeometry 60.17 $nm.min_alt_km $nm.max_inc_deg
            $visStr = if ($gNgso.isVisible) { "VISIBLE ($($gNgso.geomNote))" } else { "BELOW_HORIZON ($($gNgso.geomNote))" }
            $isVisFin = $gNgso.isVisible
        } elseif ($nm.ntc_type -eq "S") {
            $visStr = "TERRESTRIAL_EARTH_STATION"
        }

        $disposition = if ($nm.ntc_type -eq "S") {
            "EARTH_STATION_COORDINATION_CHECK"
        } elseif ($ns.totalDownlinkGrps -eq 0) {
            "UPLINK_ONLY_NO_SPACE_TO_EARTH_DOWNLINK"
        } elseif ($ns.overlapDownGrps -eq 0) {
            "CLEAR_NO_FINNISH_STATION_FREQ_OVERLAP"
        } elseif (-not $isVisFin) {
            "CLEAR_BELOW_HORIZON_IN_FINLAND"
        } elseif ($ns.exceedStationsM4.Count -gt 0 -or $ns.exceedStationsItu.Count -gt 0) {
            "CRITICAL_EXCEEDANCE_OBJECTION_REQUIRED"
        } elseif ($ns.concernStations.Count -gt 0) {
            "POTENTIAL_CONCERN_WITHIN_6DB_MONITOR"
        } elseif ($ns.wifiBtStations.Count -gt 0) {
            "LOW_SEVERITY_WIFI_BT_BAND_ONLY"
        } elseif ($null -eq $ns.worstPfd4kHz -and $ns.missingPowerGrps -gt 0) {
            "ASSESSMENT_INCOMPLETE_MISSING_FILED_POWER"
        } else {
            "COMPLIANT_WITHIN_PFD_LIMITS"
        }

        $freqRangeStr = if ($null -ne $ns.minFreqMHz) { "$([Math]::Round($ns.minFreqMHz,2)) - $([Math]::Round($ns.maxFreqMHz,2)) MHz" } else { "N/A" }

        $rowArr = @(
            $ificCode,
            $nm.ntc_id,
            $nm.sat_name,
            $nm.adm,
            $nm.ntc_type,
            $nm.prov,
            $nm.ntf_rsn,
            $nm.act_code,
            ($nm.ssn_list -join "; "),
            $orbitDesc,
            $visStr,
            $ns.totalDownlinkGrps,
            $ns.totalUplinkGrps,
            $ns.overlapDownGrps,
            $freqRangeStr,
            $ns.overlapStations.Count,
            $ns.maxGainDbi,
            $ns.maxPepDbw,
            $ns.maxPsdDbwHz,
            $ns.maxEirp4kHzDbw,
            $ns.worstPfd4kHz,
            $ns.worstMarginM4,
            $ns.worstMarginM15,
            $ns.worstMarginItu21,
            $ns.exceedStationsM4.Count,
            $ns.exceedStationsItu.Count,
            $ns.concernStations.Count,
            $ns.wifiBtStations.Count,
            (($ns.licenceClasses | Sort-Object) -join ", "),
            (($ns.bandOwners | Sort-Object) -join "; "),
            $disposition
        )
        $sheet1Rows.Add($rowArr)

        $noticeSummaryObjects.Add([PSCustomObject]@{
            ificCode = $ificCode
            ntc_id = $nm.ntc_id
            sat_name = $nm.sat_name
            adm = $nm.adm
            ntc_type = $nm.ntc_type
            prov = $nm.prov
            ssn_refs = ($nm.ssn_list -join "; ")
            orbitDesc = $orbitDesc
            long_nom = $nm.long_nom
            min_alt_km = $nm.min_alt_km
            max_inc_deg = $nm.max_inc_deg
            isVisFin = $isVisFin
            overlapStationsCount = $ns.overlapStations.Count
            exceedM4Count = $ns.exceedStationsM4.Count
            exceedItuCount = $ns.exceedStationsItu.Count
            concernCount = $ns.concernStations.Count
            worstPfd4kHz = $ns.worstPfd4kHz
            worstMarginM4 = $ns.worstMarginM4
            worstMarginItu21 = $ns.worstMarginItu21
            freqRangeStr = $freqRangeStr
            disposition = $disposition
        })
    }

    # C. Build Sheet 2: 2_Finnish_Stations_874_Impact
    $sheet2Headers = @(
        "Station_ID", "Licence_Class", "Latitude_N", "Longitude_E", "Ground_Height_m", "Mast_Height_m",
        "RX_Freq_MHz", "TX_Freq_MHz", "Channel_BW_MHz", "Ant_Gain_dBi", "Station_Azimuth_Deg", "Polarisation",
        "Taajuusjakotaulukko_Band", "Allocated_Services", "Stakeholder_Band_Owner", "Is_WiFi_BT_Band",
        "Overlapping_IFIC_Notices", "Visible_Overlapping_Notices", "Notices_Exceeding_M4", "Notices_Exceeding_ITU21", "Notices_Concern",
        "Worst_Notice_ID", "Worst_Satellite_Name", "Worst_Admin", "Worst_Orbit",
        "Elev_Deg_to_Worst_Sat", "Azimuth_Deg_to_Worst_Sat", "Off_Axis_Angle_Deg", "Slant_Range_km",
        "Worst_Arriving_PFD_dBW_m2_4kHz", "ITU_Art21_Limit_dBW_m2_4kHz", "Worst_Margin_M4_dB", "Worst_Margin_ITU21_dB",
        "Overlapping_Satellites_List", "Station_Screening_Status"
    )
    $sheet2Rows = New-Object System.Collections.Generic.List[object]
    foreach ($s in $stations) {
        $stImp = $stationImpacts[$s.idx]
        $wh = $stImp.worstHit
        $stStatus = if ($s.primaryFreqMHz -le 0) {
            "QUARANTINED_MISSING_FREQ_IN_EXCEL"
        } elseif ($stImp.overlapNotices.Count -eq 0) {
            "CLEAR_NO_SPECTRAL_OVERLAP"
        } elseif ($stImp.visibleNotices.Count -eq 0) {
            "CLEAR_OVERLAPS_BELOW_HORIZON_ONLY"
        } elseif ($null -ne $wh) {
            $wh.status
        } else {
            "COMPLIANT"
        }

        $satListStr = (($stImp.allSatNames | Select-Object -First 15) -join "; ")
        if ($stImp.allSatNames.Count -gt 15) { $satListStr += " (+$($stImp.allSatNames.Count - 15) more)" }

        $sheet2Rows.Add(@(
            $s.stationId,
            $s.licenceClass,
            $s.lat,
            $s.lon,
            $s.groundHeightM,
            $s.mastHeightM,
            $s.rxFreqMHz,
            $s.txFreqMHz,
            $s.bwMHz,
            $s.antGainDbi,
            $s.azimuthDeg,
            $s.pol,
            $s.bandInfo.bandStr,
            $s.bandInfo.services,
            $s.bandInfo.owner,
            $(if ($s.bandInfo.isWifiBt) { "YES" } else { "NO" }),
            $stImp.overlapNotices.Count,
            $stImp.visibleNotices.Count,
            $stImp.exceedNoticesM4.Count,
            $stImp.exceedNoticesItu21.Count,
            $stImp.concernNotices.Count,
            $(if ($wh) { $wh.ntc_id } else { $null }),
            $(if ($wh) { $wh.sat_name } else { "" }),
            $(if ($wh) { $wh.adm } else { "" }),
            $(if ($wh) { $wh.orbit_desc } else { "" }),
            $(if ($wh) { $wh.elev_deg } else { $null }),
            $(if ($wh) { $wh.azim_deg } else { $null }),
            $(if ($wh) { $wh.off_axis_deg } else { $null }),
            $(if ($wh) { $wh.slant_km } else { $null }),
            $(if ($wh) { $wh.pfd_4khz } else { $null }),
            $(if ($wh) { $wh.itu21_limit } else { $null }),
            $(if ($wh) { $wh.margin_m4 } else { $null }),
            $(if ($wh) { $wh.margin_itu21 } else { $null }),
            $satListStr,
            $stStatus
        ))
    }

    # D. Build Sheet 3: 3_Carrier_Station_Link_Hits
    # Sort hits by worst M4 margin descending (visible exceedances first) and take top 15,000 for rapid Excel opening while keeping full detail
    $sortedHits = $carrierHitsList | Sort-Object {
        if ($null -ne $_.margin_m4) { -$_.margin_m4 } elseif ($_.elev_deg -gt 0) { 500.0 } else { 1000.0 }
    }
    $sheet3Headers = @(
        "IFIC_Circular", "Notice_ID", "Satellite_Name", "Admin", "Orbit_Type", "Orbit_Details",
        "Provision", "Special_Section", "Beam_Name",
        "Sat_Freq_Min_MHz", "Sat_Freq_Max_MHz", "Sat_BW_MHz",
        "Beam_Gain_dBi", "PEP_Max_dBW", "PSD_Max_dBW_Hz", "EIRP_4kHz_dBW", "Power_Calculation_Method",
        "Finnish_Station_ID", "Licence_Class", "Station_Lat_N", "Station_Lon_E",
        "Station_Freq_MHz", "Station_BW_MHz", "Matched_Role", "Spectral_Overlap_MHz",
        "Taajuusjakotaulukko_Band", "Stakeholder_Band_Owner", "Is_WiFi_BT_Band",
        "Elevation_Angle_Deg", "Azimuth_Clockwise_From_North_Deg", "Station_MainBeam_Azimuth_Deg", "Off_Axis_Angle_Deg", "Slant_Range_km",
        "Arriving_PFD_PSD_dBW_m2_4kHz", "Arriving_PFD_Uniform_dBW_m2_4kHz",
        "ITU_Art21_Dynamic_Limit_dBW_m2_4kHz", "ITU_Art21_Band_Applicability_Rule",
        "Margin_vs_ITU21_dB", "Margin_vs_M4_dB (-144)", "Margin_vs_M15_dB (-154)", "Link_Screening_Status"
    )
    $sheet3Rows = New-Object System.Collections.Generic.List[object]
    $maxHitRows = [Math]::Min($sortedHits.Count, 12000)
    for ($i = 0; $i -lt $maxHitRows; $i++) {
        $h = $sortedHits[$i]
        $sheet3Rows.Add(@(
            $h.ificCode, $h.ntc_id, $h.sat_name, $h.adm, $h.ntc_type, $h.orbit_desc,
            $h.prov, $h.ssn_refs, $h.beam_name,
            $h.sat_freq_min, $h.sat_freq_max, $h.sat_bw_mhz,
            $h.gain_dbi, $h.pep_max_dbw, $h.psd_max_dbw_hz, $h.eirp_4khz_dbw, $h.power_method,
            $h.station_id, $h.licence_class, $h.lat, $h.lon,
            $h.stn_freq_mhz, $h.stn_bw_mhz, $h.matched_role, $h.overlap_mhz,
            $h.band_str, $h.band_owner, $h.is_wifi_bt,
            $h.elev_deg, $h.azim_deg, $h.stn_azim_deg, $h.off_axis_deg, $h.slant_km,
            $h.pfd_4khz, $h.pfd_uni_4khz,
            $h.itu21_limit, $h.itu21_rule,
            $h.margin_itu21, $h.margin_m4, $h.margin_m15, $h.status
        ))
    }

    # E. Build Sheet 4: 4_Band_And_Stakeholder_Summary
    $sheet4Headers = @(
        "Taajuusjakotaulukko_Band", "Allocated_Services", "Stakeholder_Band_Owner", "Is_WiFi_BT_Band",
        "Finnish_Stations_In_Band", "Stations_With_IFIC_Overlap", "Stations_With_Visible_Overlap",
        "Stations_Exceeding_M4", "Stations_Exceeding_ITU21", "Stations_Potential_Concern",
        "Worst_Arriving_PFD_dBW_m2_4kHz", "Worst_Margin_M4_dB", "Worst_Interfering_Satellite", "Overlapping_Satellites_In_Band"
    )
    $sheet4Rows = New-Object System.Collections.Generic.List[object]
    $bandGroups = $stations | Group-Object { $_.bandInfo.bandStr } | Sort-Object { $_.Group[0].primaryFreqMHz }
    foreach ($bg in $bandGroups) {
        $bSample = $bg.Group[0].bandInfo
        $stnCount = $bg.Count
        $ovCount = 0; $visCount = 0; $exM4Count = 0; $exItuCount = 0; $conCount = 0
        $worstPfd = $null; $worstM4 = $null; $worstSat = ""; $bandSats = New-Object System.Collections.Generic.HashSet[string]
        foreach ($st in $bg.Group) {
            $imp = $stationImpacts[$st.idx]
            if ($imp.overlapNotices.Count -gt 0) { $ovCount++ }
            if ($imp.visibleNotices.Count -gt 0) { $visCount++ }
            if ($imp.exceedNoticesM4.Count -gt 0) { $exM4Count++ }
            if ($imp.exceedNoticesItu21.Count -gt 0) { $exItuCount++ }
            if ($imp.concernNotices.Count -gt 0) { $conCount++ }
            foreach ($sn in $imp.allSatNames) { [void]$bandSats.Add($sn) }
            if ($null -ne $imp.worstHit -and $null -ne $imp.worstHit.pfd_4khz) {
                if ($null -eq $worstPfd -or $imp.worstHit.pfd_4khz -gt $worstPfd) {
                    $worstPfd = $imp.worstHit.pfd_4khz
                    $worstM4  = $imp.worstHit.margin_m4
                    $worstSat = "$($imp.worstHit.sat_name) ($($imp.worstHit.adm), $($imp.worstHit.orbit_desc))"
                }
            }
        }
        $sheet4Rows.Add(@(
            $bSample.bandStr, $bSample.services, $bSample.owner, $(if ($bSample.isWifiBt) { "YES" } else { "NO" }),
            $stnCount, $ovCount, $visCount, $exM4Count, $exItuCount, $conCount,
            $worstPfd, $worstM4, $worstSat, (($bandSats | Select-Object -First 12) -join "; ")
        ))
    }

    # F. Build Sheet 5: 5_Earth_Stations_And_Method (Earth Station Distance Check + Audit Notes)
    $sheet5Headers = @(
        "Section_or_Notice_ID", "Station_or_Topic_Name", "Admin_or_Category", "Coordinates_or_Formula",
        "Distance_to_Finnish_Border_km", "Freq_Range_MHz", "Evaluation_or_Audit_Resolution_Details"
    )
    $sheet5Rows = New-Object System.Collections.Generic.List[object]

    # Evaluate Earth Stations (ntc_type = 'S') against Finnish border points
    $finBorderPts = @(
        @{ name="Vaalimaa (SE Border)"; lat=60.60; lon=27.80 },
        @{ name="Imatra (SE Border)"; lat=61.17; lon=28.77 },
        @{ name="Ilomantsi (E Border)"; lat=62.67; lon=31.27 },
        @{ name="Kuusamo (NE Border)"; lat=65.96; lon=29.98 },
        @{ name="Raja-Jooseppi (Lapland E)"; lat=68.47; lon=28.47 },
        @{ name="Utsjoki (Far North)"; lat=69.90; lon=27.02 },
        @{ name="Hankoniemi (S Coast)"; lat=59.82; lon=22.97 }
    )
    foreach ($nm in ($noticeMap.Values | Where-Object { $_.ntc_type -eq "S" })) {
        $ns = $noticeStats[$nm.ntc_id.ToString()]
        $minDistKm = 99999.0
        $closestPt = ""
        if ($null -ne $nm.estn_lat -and $null -ne $nm.estn_lon) {
            foreach ($bp in $finBorderPts) {
                $dLat = ($bp.lat - $nm.estn_lat) * [Math]::PI / 180.0
                $dLon = ($bp.lon - $nm.estn_lon) * [Math]::PI / 180.0
                $a = [Math]::Pow([Math]::Sin($dLat/2),2) + [Math]::Cos($nm.estn_lat*[Math]::PI/180.0)*[Math]::Cos($bp.lat*[Math]::PI/180.0)*[Math]::Pow([Math]::Sin($dLon/2),2)
                $dKm = 2.0 * 6371.0 * [Math]::Asin([Math]::Min(1.0, [Math]::Sqrt($a)))
                if ($dKm -lt $minDistKm) { $minDistKm = [Math]::Round($dKm, 1); $closestPt = $bp.name }
            }
        }
        $fStr = if ($null -ne $ns.minFreqMHz) { "$([Math]::Round($ns.minFreqMHz,2)) - $([Math]::Round($ns.maxFreqMHz,2)) MHz" } else { "N/A" }
        $evalNote = if ($minDistKm -lt 800.0) {
            "WITHIN COORDINATION CONTOUR (<800 km of $closestPt): Check ITU RR Appendix 7 earth station horizon gain and tropospheric/ducting propagation."
        } else {
            "BEYOND TERRESTRIAL COORDINATION DISTANCE ($minDistKm km from $closestPt): No direct overland interference to Finland."
        }
        $sheet5Rows.Add(@(
            "EARTH_STN: $($nm.ntc_id)", $nm.sat_name, $nm.adm, "$($nm.estn_lat)°N, $($nm.estn_lon)°E",
            $minDistKm, $fStr, $evalNote
        ))
    }

    # Add Methodology & Claude Audit Resolution Notes
    $auditNotes = @(
        @("AUDIT_FIX_1", "Power Spectral Density vs Uniform Model", "ITU SRS emiss.pwr_ds_max + s_beam.gain", "EIRP_4kHz = pwr_ds_max + gain + 10*log10(4000)", 0, "All Bands", "Uses actual filed ITU maximum power spectral density (pwr_ds_max in dBW/Hz) plus satellite beam peak gain (s_beam.gain in dBi), capped by total peak EIRP (pep_max + gain). Also outputs uniform-spectrum PFD for comparison."),
        @("AUDIT_FIX_2", "ITU RR Table 21-4 Band Applicability Gate", "Explicit <1525 MHz Gate", "hasTable21Limit = (Freq >= 1525 MHz)", 0, "<1525 MHz vs >=1525 MHz", "Bands below 1525 MHz (e.g. VHF/UHF) do not have an ITU RR Table 21-4 PFD limit; they are explicitly flagged as COORD_TRIGGER_BAND rather than applying an out-of-band Table 21-4 limit, while retaining Traficom M4 (-144 dBW/m²/4kHz) national screening."),
        @("AUDIT_FIX_3", "GSO vs Non-GSO 3D Orbital Geometry", "Exact Spherical + NGSO Orbit Table", "GSO: R_sat=42157km; NGSO: orbit(perig_km, inclin_ang)", 0, "GSO & NGSO", "Evaluates exact topocentric elevation & slant range for GSO slots (flagging BELOW_HORIZON when delta <= 0 deg in Finland) and uses filed perigee altitude & inclination for Non-GSO constellations."),
        @("AUDIT_FIX_4", "Corrected Clockwise Azimuth & Off-Axis Angle", "Clockwise from True North (0..360 deg)", "atan2(sin(dLon), -sin(lat)*cos(dLon))", 0, "All Links", "Fixed North-South sign inversion so a geostationary satellite at 25E observed from Finland (60N-70N, 25E) yields ~180 deg (due South), and computes off-axis angle against Finnish station azimuths."),
        @("AUDIT_FIX_5", "Missing Power / Incomplete Records Gate", "Explicit Quarantine Status", "ASSESSMENT_INCOMPLETE / QUARANTINED", 0, "Data Hygiene", "Quarantines the 3 incomplete 0-MHz Finnish station rows (362450, 368852, 368856) and explicitly marks any satellite group lacking power/gain data as ASSESSMENT_INCOMPLETE_MISSING_POWER rather than Compliant.")
    )
    foreach ($an in $auditNotes) { $sheet5Rows.Add($an) }

    # Write Multi-Sheet .xlsx and clean .csv files
    $xlsxPath = Join-Path $outDir "$($ificCode)_Full_Predictions.xlsx"
    Write-Output "   Writing $xlsxPath ..."
    Write-MultiSheetXlsx $xlsxPath @(
        @{ Name = "1_Executive_Summary_By_Notice"; Headers = $sheet1Headers; Rows = $sheet1Rows },
        @{ Name = "2_Finnish_Stations_874_Impact"; Headers = $sheet2Headers; Rows = $sheet2Rows },
        @{ Name = "3_Carrier_Station_Link_Hits";   Headers = $sheet3Headers; Rows = $sheet3Rows },
        @{ Name = "4_Band_And_Stakeholder_Summary";Headers = $sheet4Headers; Rows = $sheet4Rows },
        @{ Name = "5_Earth_Stations_And_Method";   Headers = $sheet5Headers; Rows = $sheet5Rows }
    )

    Write-CleanCsv (Join-Path $outDir "$($ificCode)_Notice_Executive_Summary.csv") $sheet1Headers $sheet1Rows
    Write-CleanCsv (Join-Path $outDir "$($ificCode)_Finnish_Stations_874_Impact.csv") $sheet2Headers $sheet2Rows
    Write-CleanCsv (Join-Path $outDir "$($ificCode)_Carrier_Station_Link_Hits.csv") $sheet3Headers $sheet3Rows

        $circDuration = $circStopwatch.Elapsed
    $circSecs = [Math]::Round($circDuration.TotalSeconds, 2)
    Write-Output "   Finished $ificTitle in ${circSecs}s ($($circDuration.ToString('mm\:ss')))"
    return [PSCustomObject]@{
        IficCode       = $ificCode
        Sheet1Headers  = $sheet1Headers
        Sheet1Rows     = $sheet1Rows
        Sheet2Headers  = $sheet2Headers
        Sheet2Rows     = $sheet2Rows
        Sheet4Headers  = $sheet4Headers
        Sheet4Rows     = $sheet4Rows
        NoticeObjects  = $noticeSummaryObjects
        TopHits        = ($sortedHits | Select-Object -First 200)
    }
}

$res3074 = $null
$res3076 = $null

foreach ($job in $targetMdbJobs) {
    $res = Invoke-IficScreening $job.MdbPath $job.IficCode $job.IficTitle
    if ($job.IficCode -eq "BR_IFIC_3074") { $res3074 = $res }
    if ($job.IficCode -eq "BR_IFIC_3076") { $res3076 = $res }
}

# Write Combined Executive Summary Workbook if both 3074 and 3076 were processed
if ($runDefaultPair -and $null -ne $res3074 -and $null -ne $res3076) {
    $combinedXlsx = Join-Path $outDir "BR_IFIC_3074_3076_Combined_Executive_Summary.xlsx"
    $combinedNoticeRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $res3074.Sheet1Rows) { $combinedNoticeRows.Add($r) }
    foreach ($r in $res3076.Sheet1Rows) { $combinedNoticeRows.Add($r) }

    Write-MultiSheetXlsx $combinedXlsx @(
        @{ Name = "All_Notices_3074_And_3076";      Headers = $res3074.Sheet1Headers; Rows = $combinedNoticeRows },
        @{ Name = "IFIC_3074_Station_Impacts";      Headers = $res3074.Sheet2Headers; Rows = $res3074.Sheet2Rows },
        @{ Name = "IFIC_3076_Station_Impacts";      Headers = $res3076.Sheet2Headers; Rows = $res3076.Sheet2Rows },
        @{ Name = "IFIC_3074_Band_Summary";         Headers = $res3074.Sheet4Headers; Rows = $res3074.Sheet4Rows },
        @{ Name = "IFIC_3076_Band_Summary";         Headers = $res3076.Sheet4Headers; Rows = $res3076.Sheet4Rows }
    )
}

Write-Output "All Excel (.xlsx) and CSV files successfully created in $outDir !"

$totalPipelineStopwatch.Stop()
$pipelineEndTime = Get-Date
$elapsed = $totalPipelineStopwatch.Elapsed
$formattedElapsed = if ($elapsed.TotalHours -ge 1) {
    "{0:D2}h {1:D2}m {2:D2}s ({3:F1}s)" -f [int]$elapsed.TotalHours, $elapsed.Minutes, $elapsed.Seconds, $elapsed.TotalSeconds
} elseif ($elapsed.TotalMinutes -ge 1) {
    "{0:D2}m {1:D2}s ({2:F1}s)" -f [int]$elapsed.TotalMinutes, $elapsed.Seconds, $elapsed.TotalSeconds
} else {
    "{0:F2} seconds" -f $elapsed.TotalSeconds
}

Write-Output ""
Write-Output "=========================================================="
Write-Output "SCRIPT EXECUTION TIMER SUMMARY"
Write-Output "Start Time:     $($pipelineStartTime.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture))"
Write-Output "Stop Time:      $($pipelineEndTime.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture))"
Write-Output "Total Runtime:  $formattedElapsed"
Write-Output "Output Folder:  $outDir"
Write-Output "=========================================================="
