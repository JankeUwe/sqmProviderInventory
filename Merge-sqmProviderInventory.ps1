<#
.SYNOPSIS
    Merges the per-server JSON files of Get-sqmProviderInventory.ps1 into Excel workbooks.

.DESCRIPTION
    Reads all scan files of a folder (the newest file per server wins), applies the mapping CSV
    (application, parent group, protection need, ...) and writes:

      Provider_Inventory_<date>.xlsx   For the new provider. NO logins, users or roles inside.
      Security_Inventory_<date>.xlsx   Only with -IncludeSecurity (step 2). Not for the provider.
      Pruefbericht_<date>.txt          Check report: missing scans, missing mapping, errors.

    Needs no Excel and no extra module. The workbooks are written directly (OpenXML).
    Runs on Windows PowerShell 5.1 and PowerShell 7.

.PARAMETER InputPath
    Folder with the JSON files (or a list of files).

.PARAMETER MappingPath
    Mapping CSV, delimiter ";", UTF-8. Columns: Server;Anwendung;Obergruppe;Schutzbedarf;
    Kritikalitaet;RPO;RTO;Wartungsfenster;Ansprechpartner;Bemerkung.
    Several applications per server: separate with "|". The CSV wins over values set at scan time.

.PARAMETER NewMappingTemplate
    Writes (or extends) the mapping CSV with all scanned servers and exits. Existing rows and
    their values are kept.

.PARAMETER OutputPath
    Target folder. Default: InputPath.

.PARAMETER IncludeSecurity
    Also write the security workbook (logins, database users, roles, permissions).

.PARAMETER IncludeSystemAndSysadmin
    Security workbook: do not filter system logins and sysadmin members (default: filtered).

.PARAMETER Anonymize
    Replaces server names, instance names and the domain in the provider workbook by neutral IDs
    (SRV001 ...). The key is written to Anonymize_Zuordnung_<date>.csv. Do NOT pass that file on.

.PARAMETER MaxAgeDays
    Scans older than this are reported in the check report. Default 30.

.EXAMPLE
    .\Merge-sqmProviderInventory.ps1 -InputPath D:\Inventory -MappingPath D:\Inventory\Mapping.csv -NewMappingTemplate

.EXAMPLE
    .\Merge-sqmProviderInventory.ps1 -InputPath D:\Inventory -MappingPath D:\Inventory\Mapping.csv -IncludeSecurity
#>
[CmdletBinding()]
param (
	[Parameter(Mandatory = $true)]
	[string[]]$InputPath,
	[string]$MappingPath,
	[switch]$NewMappingTemplate,
	[string]$OutputPath,
	[switch]$IncludeSecurity,
	[switch]$IncludeSystemAndSysadmin,
	[switch]$Anonymize,
	[int]$MaxAgeDays = 30
)

$ErrorActionPreference = 'Stop'
$stamp = Get-Date -Format 'yyyyMMdd'
$MappingColumns = @('Server', 'Anwendung', 'Obergruppe', 'Schutzbedarf', 'Kritikalitaet', 'RPO', 'RTO', 'Wartungsfenster', 'Ansprechpartner', 'Bemerkung')

#region xlsx writer
Add-Type -AssemblyName System.IO.Compression
$script:XmlBad = New-Object System.Text.RegularExpressions.Regex('[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]')

function ConvertTo-ColumnName
{
	param ([int]$Index)
	$s = ''
	$n = $Index + 1
	while ($n -gt 0)
	{
		$m = [int](($n - 1) % 26)
		$s = [string][char](65 + $m) + $s
		$n = [int][math]::Floor(($n - 1) / 26)
	}
	return $s
}

function Get-SafeSheetName
{
	param ([string]$Name, $Used)
	$n = ($Name -replace '[\\/\?\*\[\]:]', '_')
	if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
	if (-not $n) { $n = 'Sheet' }
	$base = $n
	$i = 2
	while ($Used -contains $n.ToLowerInvariant())
	{
		$suffix = "_$i"
		$n = $base.Substring(0, [math]::Min($base.Length, 31 - $suffix.Length)) + $suffix
		$i++
	}
	return $n
}

function Write-XlsxFile
{
	param ([string]$Path, [System.Collections.Specialized.OrderedDictionary]$Sheets)

	$inv = [System.Globalization.CultureInfo]::InvariantCulture
	if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
	$fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Create)
	$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
	$enc = New-Object System.Text.UTF8Encoding($false)
	try
	{
		$writeEntry = {
			param ($name, $text)
			$e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
			$sw = New-Object System.IO.StreamWriter($e.Open(), $enc)
			$sw.Write($text)
			$sw.Dispose()
		}

		$used = @()
		$sheetNames = @()
		foreach ($k in $Sheets.Keys)
		{
			$nm = Get-SafeSheetName -Name $k -Used $used
			$used += $nm.ToLowerInvariant()
			$sheetNames += $nm
		}

		$ct = New-Object System.Text.StringBuilder
		[void]$ct.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
		for ($i = 1; $i -le $sheetNames.Count; $i++) { [void]$ct.Append("<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>") }
		[void]$ct.Append('</Types>')
		& $writeEntry '[Content_Types].xml' $ct.ToString()
		& $writeEntry '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

		$wb = New-Object System.Text.StringBuilder
		[void]$wb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')
		$rel = New-Object System.Text.StringBuilder
		[void]$rel.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
		for ($i = 1; $i -le $sheetNames.Count; $i++)
		{
			$esc = [System.Security.SecurityElement]::Escape($sheetNames[$i - 1])
			[void]$wb.Append("<sheet name=`"$esc`" sheetId=`"$i`" r:id=`"rId$i`"/>")
			[void]$rel.Append("<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>")
		}
		[void]$wb.Append('</sheets></workbook>')
		$sid = $sheetNames.Count + 1
		[void]$rel.Append("<Relationship Id=`"rId$sid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>")
		& $writeEntry 'xl/workbook.xml' $wb.ToString()
		& $writeEntry 'xl/_rels/workbook.xml.rels' $rel.ToString()
		& $writeEntry 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs></styleSheet>'

		$idx = 0
		foreach ($key in $Sheets.Keys)
		{
			$idx++
			$rows = @($Sheets[$key])
			$cols = New-Object System.Collections.ArrayList
			$seen = @{ }
			foreach ($r in $rows)
			{
				foreach ($p in $r.PSObject.Properties) { if (-not $seen.ContainsKey($p.Name)) { $seen[$p.Name] = $true; [void]$cols.Add($p.Name) } }
			}
			if ($cols.Count -eq 0) { [void]$cols.Add('(keine Daten)') }
			$widths = New-Object 'int[]' $cols.Count
			for ($c = 0; $c -lt $cols.Count; $c++) { $widths[$c] = [math]::Min([math]::Max($cols[$c].Length + 2, 8), 60) }

			$sb = New-Object System.Text.StringBuilder
			[void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
			$lastCol = ConvertTo-ColumnName ($cols.Count - 1)
			$lastRow = $rows.Count + 1
			[void]$sb.Append("<dimension ref=`"A1:$lastCol$lastRow`"/>")
			[void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="13"/>')
			$sheetDataSb = New-Object System.Text.StringBuilder
			[void]$sheetDataSb.Append('<sheetData><row r="1" ht="28" customHeight="1">')
			for ($c = 0; $c -lt $cols.Count; $c++)
			{
				$ref = (ConvertTo-ColumnName $c) + '1'
				$t = [System.Security.SecurityElement]::Escape($script:XmlBad.Replace([string]$cols[$c], ''))
				[void]$sheetDataSb.Append("<c r=`"$ref`" s=`"1`" t=`"inlineStr`"><is><t>$t</t></is></c>")
			}
			[void]$sheetDataSb.Append('</row>')
			$rn = 1
			foreach ($r in $rows)
			{
				$rn++
				[void]$sheetDataSb.Append("<row r=`"$rn`">")
				for ($c = 0; $c -lt $cols.Count; $c++)
				{
					$prop = $r.PSObject.Properties[$cols[$c]]
					if (-not $prop) { continue }
					$v = $prop.Value
					if ($null -eq $v) { continue }
					$ref = (ConvertTo-ColumnName $c) + $rn
					if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal] -or $v -is [single] -or $v -is [int16] -or $v -is [byte] -or $v -is [uint32] -or $v -is [uint64])
					{
						if (($v -is [double] -or $v -is [single]) -and ([double]::IsNaN([double]$v) -or [double]::IsInfinity([double]$v))) { continue }
						[void]$sheetDataSb.Append("<c r=`"$ref`"><v>" + ([System.Convert]::ToString($v, $inv)) + '</v></c>')
						if ($widths[$c] -lt 14) { $widths[$c] = 14 }
					}
					else
					{
						if ($v -is [bool]) { $s = if ($v) { 'Ja' } else { 'Nein' } }
						elseif ($v -is [datetime]) { $s = $v.ToString('yyyy-MM-ddTHH:mm:ss') }
						elseif ($v -is [System.Array]) { $s = ($v | ForEach-Object { [string]$_ }) -join '; ' }
						else { $s = [string]$v }
						if ($s.Length -gt 32000) { $s = $s.Substring(0, 32000) }
						$s = $script:XmlBad.Replace($s, '')
						if ($s.Length -eq 0) { continue }
						$esc = [System.Security.SecurityElement]::Escape($s)
						[void]$sheetDataSb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$esc</t></is></c>")
						$w = [math]::Min($s.Length + 2, 60)
						if ($w -gt $widths[$c]) { $widths[$c] = $w }
					}
				}
				[void]$sheetDataSb.Append('</row>')
			}
			[void]$sheetDataSb.Append('</sheetData>')
			[void]$sb.Append('<cols>')
			for ($c = 0; $c -lt $cols.Count; $c++) { [void]$sb.Append("<col min=`"$($c + 1)`" max=`"$($c + 1)`" width=`"$($widths[$c])`" customWidth=`"1`"/>") }
			[void]$sb.Append('</cols>')
			[void]$sb.Append($sheetDataSb.ToString())
			[void]$sb.Append("<autoFilter ref=`"A1:$lastCol$lastRow`"/></worksheet>")
			& $writeEntry "xl/worksheets/sheet$idx.xml" $sb.ToString()
		}
	}
	finally
	{
		$zip.Dispose()
		$fs.Dispose()
	}
}
#endregion

#region json reading (5.1 safe)
function ConvertTo-PlainObject
{
	param ($Node)
	if ($null -eq $Node) { return $null }
	if ($Node -is [System.Collections.IDictionary])
	{
		$o = [ordered]@{ }
		foreach ($k in $Node.Keys) { $o[[string]$k] = ConvertTo-PlainObject $Node[$k] }
		return [pscustomobject]$o
	}
	if ($Node -is [System.Collections.IList] -and $Node -isnot [string])
	{
		$list = New-Object System.Collections.Generic.List[object]
		foreach ($i in $Node) { $list.Add((ConvertTo-PlainObject $i)) }
		return , $list.ToArray()
	}
	return $Node
}

function Read-ScanFile
{
	param ([string]$Path)
	$text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
	if ($PSVersionTable.PSEdition -eq 'Core')
	{
		# keep ISO dates as text (PowerShell 7.5+ has -DateKind, older versions are handled by the xlsx writer)
		try { return ($text | ConvertFrom-Json -DateKind String) } catch { return ($text | ConvertFrom-Json) }
	}
	Add-Type -AssemblyName System.Web.Extensions
	$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
	$ser.MaxJsonLength = [int]::MaxValue
	$ser.RecursionLimit = 200
	return (ConvertTo-PlainObject ($ser.DeserializeObject($text)))
}
#endregion

#region helpers
function Get-Prop
{
	param ($Object, [string]$Name)
	if ($null -eq $Object) { return $null }
	$p = $Object.PSObject.Properties[$Name]
	if ($p) { return $p.Value }
	return $null
}
# @($null) is a one-element array; this returns a real empty array for missing properties
function AsList
{
	param ($V)
	if ($null -eq $V) { return @() }
	return @($V | Where-Object { $null -ne $_ })
}
function Round-Or
{
	param ($V, [int]$D = 1)
	if ($null -eq $V -or $V -eq '') { return $null }
	return [math]::Round([double]$V, $D)
}
function ToGb
{
	param ($Mb, [int]$D = 1)
	if ($null -eq $Mb -or $Mb -eq '') { return $null }
	return [math]::Round([double]$Mb / 1024, $D)
}
function ToMin
{
	param ($Sec, [int]$D = 1)
	if ($null -eq $Sec -or $Sec -eq '') { return $null }
	return [math]::Round([double]$Sec / 60, $D)
}
function Join-Unique
{
	param ($Items)
	return ((@($Items) | Where-Object { $_ } | Select-Object -Unique) -join '; ')
}
function Add-Row
{
	param ($List, $Row)
	[void]$List.Add([pscustomobject]$Row)
}
#endregion

#region load
$files = @()
foreach ($p in $InputPath)
{
	if (Test-Path -LiteralPath $p -PathType Container) { $files += @(Get-ChildItem -LiteralPath $p -Filter '*.json' -File) }
	elseif (Test-Path -LiteralPath $p -PathType Leaf) { $files += Get-Item -LiteralPath $p }
	else { throw "Input path not found: $p" }
}
$files = @($files | Where-Object { $_.Name -notmatch '^(Mapping|Anonymize)' })
if ($files.Count -eq 0) { throw 'No JSON scan files found.' }
if (-not $OutputPath)
{
	$first = $InputPath[0]
	$OutputPath = if (Test-Path -LiteralPath $first -PathType Container) { $first } else { Split-Path -Parent $first }
}
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

$report = New-Object System.Collections.ArrayList
function Add-Report { param ([string]$Level, [string]$Server, [string]$Text) [void]$report.Add([pscustomobject][ordered]@{ Stufe = $Level; Server = $Server; Meldung = $Text }) }

$byHost = @{ }
foreach ($f in $files)
{
	try { $scan = Read-ScanFile -Path $f.FullName }
	catch { Add-Report 'FEHLER' $f.Name "Datei nicht lesbar: $($_.Exception.Message)"; continue }
	$hn = [string](Get-Prop (Get-Prop $scan 'host') 'computer_name')
	if (-not $hn) { Add-Report 'FEHLER' $f.Name 'Kein host.computer_name, Datei ist kein gueltiger Scan.'; continue }
	$key = $hn.ToUpperInvariant()
	$scan | Add-Member -NotePropertyName _file -NotePropertyValue $f.Name -Force
	if ($byHost.ContainsKey($key))
	{
		$old = $byHost[$key]
		if ([datetime]$scan.scan_started -gt [datetime]$old.scan_started)
		{
			Add-Report 'HINWEIS' $hn "Mehrere Scans, aelterer ignoriert: $($old._file)"
			$byHost[$key] = $scan
		}
		else { Add-Report 'HINWEIS' $hn "Mehrere Scans, aelterer ignoriert: $($f.Name)" }
	}
	else { $byHost[$key] = $scan }
}
$scans = @($byHost.Values | Sort-Object { $_.host.computer_name })

# mapping
$mapping = @{ }
$mappingRows = @()
if ($MappingPath -and (Test-Path -LiteralPath $MappingPath))
{
	$mappingRows = @(Import-Csv -LiteralPath $MappingPath -Delimiter ';' -Encoding UTF8)
	foreach ($m in $mappingRows) { if ($m.Server) { $mapping[$m.Server.Trim().ToUpperInvariant()] = $m } }
}

if ($NewMappingTemplate)
{
	if (-not $MappingPath) { $MappingPath = Join-Path $OutputPath 'Mapping.csv' }
	$rows = New-Object System.Collections.ArrayList
	$known = @{ }
	foreach ($m in $mappingRows)
	{
		$o = [ordered]@{ }
		foreach ($c in $MappingColumns) { $o[$c] = [string](Get-Prop $m $c) }
		[void]$rows.Add([pscustomobject]$o)
		$known[$m.Server.Trim().ToUpperInvariant()] = $true
	}
	$added = 0
	foreach ($s in $scans)
	{
		$k = $s.host.computer_name.ToUpperInvariant()
		if ($known.ContainsKey($k)) { continue }
		$o = [ordered]@{ }
		foreach ($c in $MappingColumns) { $o[$c] = '' }
		$o.Server = $s.host.computer_name
		$o.Anwendung = [string](Get-Prop $s.mapping 'application')
		$o.Obergruppe = [string](Get-Prop $s.mapping 'parent_group')
		[void]$rows.Add([pscustomobject]$o)
		$added++
	}
	$rows | Export-Csv -LiteralPath $MappingPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
	Write-Host "Mapping template: $MappingPath ($added new server(s), $($rows.Count) total)"
	return
}

function Get-MapValue
{
	param ($Scan, [string]$Column, [string]$ScanKey)
	$k = $Scan.host.computer_name.ToUpperInvariant()
	$v = $null
	if ($mapping.ContainsKey($k)) { $v = [string](Get-Prop $mapping[$k] $Column) }
	if (-not $v -and $ScanKey) { $v = [string](Get-Prop $Scan.mapping $ScanKey) }
	return $v
}
#endregion

#region check report (inputs)
$now = Get-Date
foreach ($m in $mappingRows)
{
	if ($m.Server -and -not $byHost.ContainsKey($m.Server.Trim().ToUpperInvariant())) { Add-Report 'FEHLT' $m.Server 'Im Mapping, aber kein Scan vorhanden.' }
}
foreach ($s in $scans)
{
	$hn = $s.host.computer_name
	if (-not (Get-MapValue $s 'Anwendung' 'application')) { Add-Report 'ZUORDNUNG' $hn 'Anwendung fehlt (NICHT ZUGEORDNET).' }
	if (-not (Get-MapValue $s 'Obergruppe' 'parent_group')) { Add-Report 'ZUORDNUNG' $hn 'Obergruppe fehlt (NICHT ZUGEORDNET).' }
	if ($mappingRows.Count -gt 0 -and -not $mapping.ContainsKey($hn.ToUpperInvariant())) { Add-Report 'ZUORDNUNG' $hn 'Server fehlt im Mapping.' }
	$age = ($now - [datetime]$s.scan_started).TotalDays
	if ($age -gt $MaxAgeDays) { Add-Report 'ALT' $hn ("Scan ist {0:N0} Tage alt ({1})." -f $age, $s.scan_started) }
	if ([string]$s.schema_version -ne '1.0') { Add-Report 'HINWEIS' $hn "Unbekannte Schema-Version $($s.schema_version)." }
	foreach ($e in (AsList $s.host.errors)) { Add-Report 'LUECKE' $hn "Host $($e.section): $($e.message)" }
	if (@($s.instances).Count -eq 0) { Add-Report 'HINWEIS' $hn 'Keine SQL-Instanz im Scan.' }
	foreach ($i in (AsList $s.instances))
	{
		if (-not $i.connected) { Add-Report 'FEHLER' $hn "Instanz $($i.connect_target) nicht erreichbar: $($i.connect_error)"; continue }
		foreach ($e in (AsList $i.errors)) { Add-Report 'LUECKE' $hn "Instanz $($i.instance_name) $($e.section): $($e.message)" }
		if ($i.server -and $i.server.uptime_hours -lt 168) { Add-Report 'HINWEIS' $hn ("Instanz {0}: SQL-Dienst laeuft erst {1} h, Lastwerte sind nicht belastbar." -f $i.instance_name, $i.server.uptime_hours) }
	}
}
#endregion

#region build provider sheets
$S_server = New-Object System.Collections.ArrayList
$S_disks = New-Object System.Collections.ArrayList
$S_hw = New-Object System.Collections.ArrayList
$S_inst = New-Object System.Collections.ArrayList
$S_db = New-Object System.Collections.ArrayList
$S_files = New-Object System.Collections.ArrayList
$S_dep = New-Object System.Collections.ArrayList
$S_ha = New-Object System.Collections.ArrayList
$S_comp = New-Object System.Collections.ArrayList
$S_jobs = New-Object System.Collections.ArrayList
$anyJobs = $false
$reportedPartners = @{ }

foreach ($s in $scans)
{
	$h = $s.host
	$hn = $h.computer_name
	$app = Get-MapValue $s 'Anwendung' 'application'
	$grp = Get-MapValue $s 'Obergruppe' 'parent_group'
	if (-not $app) { $app = 'NICHT ZUGEORDNET' }
	if (-not $grp) { $grp = 'NICHT ZUGEORDNET' }
	$prot = Get-MapValue $s 'Schutzbedarf' $null
	$crit = Get-MapValue $s 'Kritikalitaet' $null
	$insts = @($s.instances)
	$cpu = $h.cpu

	Add-Row $S_server ([ordered]@{
			Server = $hn; Anwendung = $app; Obergruppe = $grp; Schutzbedarf = $prot; Kritikalitaet = $crit
			RPO = (Get-MapValue $s 'RPO' $null); RTO = (Get-MapValue $s 'RTO' $null); Wartungsfenster = (Get-MapValue $s 'Wartungsfenster' $null); Ansprechpartner = (Get-MapValue $s 'Ansprechpartner' $null)
			FQDN = $h.fqdn; Domaene = $h.domain; 'Domaenenmitglied' = $h.part_of_domain
			Virtuell = $h.is_virtual; Hersteller = $h.manufacturer; Modell = $h.model
			Betriebssystem = $h.os.caption; 'OS-Version' = $h.os.version; 'OS-Build' = $h.os.build; 'OS-Edition' = $h.os.edition_id; Installationsart = $h.os.installation_type; Architektur = $h.os.architecture
			'OS-Release' = $h.os.display_version; 'Letzter Hotfix' = (Get-Prop $h.last_hotfix 'installed_on')
			Sockets = $cpu.sockets; Kerne = $cpu.cores; 'Logische CPUs' = $cpu.logical_processors; 'CPU-Modell' = ([string]$cpu.model).Trim()
			'RAM GB' = (ToGb $h.memory_mb 1); 'Pagefile MB' = $h.pagefile_configured_mb; Energieplan = $h.power_plan
			'Letzter Boot' = $h.last_boot; Zeitzone = $h.time_zone
			IPs = (Join-Unique (@($h.ip_addresses) | ForEach-Object { @($_.ip) | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' } }))
			'Anzahl Instanzen' = $insts.Count; 'Scan-Zeitpunkt' = $s.scan_started; 'Scan-Dauer s' = $s.scan_duration_sec
		})

	# disks: add SQL IO per volume (sum over instances)
	foreach ($v in (AsList $h.volumes))
	{
		$sqlMb = 0.0; $iops = 0.0; $mbps = 0.0; $rl = $null; $wl = $null; $cnt = 0
		foreach ($i in $insts)
		{
			foreach ($vi in (AsList $i.volume_io))
			{
				if ($vi.volume -eq $v.name)
				{
					$sqlMb += [double]$vi.sql_files_mb; $iops += [double]$vi.avg_iops_since_start; $mbps += [double]$vi.avg_mb_per_sec_since_start; $cnt += [int]$vi.sql_file_count
					if ($null -ne $vi.avg_read_latency_ms -and ($null -eq $rl -or $vi.avg_read_latency_ms -gt $rl)) { $rl = $vi.avg_read_latency_ms }
					if ($null -ne $vi.avg_write_latency_ms -and ($null -eq $wl -or $vi.avg_write_latency_ms -gt $wl)) { $wl = $vi.avg_write_latency_ms }
				}
			}
		}
		$pct = if ($v.size_gb) { [math]::Round(100.0 * $v.used_gb / $v.size_gb, 1) } else { $null }
		Add-Row $S_disks ([ordered]@{
				Server = $hn; Obergruppe = $grp; Laufwerk = $v.name; Label = $v.label; Dateisystem = $v.file_system; 'Blockgroesse B' = $v.block_size
				'Groesse GB' = $v.size_gb; 'Belegt GB' = $v.used_gb; 'Frei GB' = $v.free_gb; 'Belegt %' = $pct
				'SQL-Dateien GB' = $(if ($cnt -gt 0) { ToGb $sqlMb 1 } else { $null }); 'SQL-Dateien Anzahl' = $(if ($cnt -gt 0) { $cnt } else { $null })
				'IOPS avg (seit SQL-Start)' = $(if ($cnt -gt 0) { Round-Or $iops 1 } else { $null }); 'MB/s avg' = $(if ($cnt -gt 0) { Round-Or $mbps 2 } else { $null })
				'Read-Latenz ms' = $rl; 'Write-Latenz ms' = $wl
			})
	}
	foreach ($pd in (AsList $h.physical_disks)) { Add-Row $S_hw ([ordered]@{ Server = $hn; Quelle = 'PhysicalDisk'; Bezeichnung = $pd.friendly_name; Typ = $pd.media_type; Bus = $pd.bus_type; 'Groesse GB' = $pd.size_gb }) }
	if (@($h.physical_disks).Count -eq 0) { foreach ($dd in (AsList $h.disk_drives)) { Add-Row $S_hw ([ordered]@{ Server = $hn; Quelle = 'DiskDrive'; Bezeichnung = $dd.model; Typ = $dd.media_type; Bus = $dd.interface; 'Groesse GB' = $dd.size_gb }) } }

	# host level components
	foreach ($sv in (AsList $h.sql_services)) { Add-Row $S_comp ([ordered]@{ Server = $hn; Instanz = $null; Typ = 'SQL-Dienst'; Name = $sv.display_name; Detail = "$($sv.state) / $($sv.start_mode)"; Konto = $sv.account }) }
	foreach ($sv in (AsList $h.third_party_services)) { Add-Row $S_comp ([ordered]@{ Server = $hn; Instanz = $null; Typ = 'Dienst (Fremdsoftware)'; Name = $sv.display_name; Detail = $sv.state; Konto = $sv.account }) }

	foreach ($i in $insts)
	{
		if (-not $i.connected) { continue }
		$sv = $i.server
		$ld = $i.load
		$dbs = @($i.databases)
		$userDbs = @($dbs | Where-Object { $_.name -ne 'tempdb' })
		$dataGb = ToGb (($userDbs | Measure-Object -Property data_size_mb -Sum).Sum) 1
		$port = $null
		$tcpFixed = (AsList $i.tcp_settings | Where-Object { $_.registry_key -match 'IPAll$' -and $_.value_name -eq 'TcpPort' } | Select-Object -First 1).value_data
		$tcpDyn = (AsList $i.tcp_settings | Where-Object { $_.registry_key -match 'IPAll$' -and $_.value_name -eq 'TcpDynamicPorts' } | Select-Object -First 1).value_data
		$port = if ($tcpFixed) { ([string]$tcpFixed).Trim([char]0, ' ') } elseif ($tcpDyn) { (([string]$tcpDyn).Trim([char]0, ' ') + ' (dynamisch)') } else { $null }
		if ($port -and -not ($port -replace '\(dynamisch\)', '').Trim()) { $port = $null }
		# Availability group info of this node: AG names, local role, partner nodes
		$agRows = AsList $i.high_availability.availability_replicas
		$agNames = Join-Unique ($agRows | ForEach-Object { $_.ag_name })
		$agLocal = Join-Unique ($agRows | Where-Object { $_.is_local -eq $true -or $_.is_local -eq 1 } | ForEach-Object { "$($_.ag_name): $($_.role_desc)" })
		$agPartners = Join-Unique ($agRows | Where-Object { -not ($_.is_local -eq $true -or $_.is_local -eq 1) } | ForEach-Object { $_.replica_server_name })
		$agDbMap = @{ }
		foreach ($ad in (AsList $i.high_availability.availability_databases)) { $agDbMap[[string]$ad.database_name] = [string]$ad.ag_name }
		foreach ($pn in (AsList $i.high_availability.availability_replicas))
		{
			$phost = ([string]$pn.replica_server_name).Split('\')[0].Split('.')[0].ToUpperInvariant()
			if ($phost -and -not $byHost.ContainsKey($phost) -and -not $reportedPartners.ContainsKey("$hn|$phost"))
			{
				$reportedPartners["$hn|$phost"] = $true
				Add-Report 'FEHLT' $phost "AG-Partnerknoten von $hn (AG $($pn.ag_name)) ist nicht gescannt."
			}
		}
		$cfg = $i.config
		Add-Row $S_inst ([ordered]@{
				Server = $hn; Anwendung = $app; Obergruppe = $grp; Instanz = $i.instance_name
				'SQL-Version' = $sv.product_version; 'Major' = $sv.major_version; Patchlevel = $sv.product_level; 'CU' = $sv.product_update_level; Edition = $sv.edition
				Collation = $sv.collation; 'Lizenz-Kerne (sichtbar)' = $sv.licensed_cores_visible_online; 'CPU gesamt' = $sv.cpu_count
				'RAM Server MB' = $sv.physical_memory_mb; 'Max Memory MB' = (Get-Prop $cfg 'max server memory (MB)'); 'Min Memory MB' = (Get-Prop $cfg 'min server memory (MB)'); MAXDOP = (Get-Prop $cfg 'max degree of parallelism')
				'Cost Threshold' = (Get-Prop $cfg 'cost threshold for parallelism'); 'Backup Compression Default' = (Get-Prop $cfg 'backup compression default')
				Cluster = [bool]$sv.is_clustered; 'AlwaysOn aktiv' = [bool]$sv.is_hadr_enabled; 'Nur Windows-Auth' = [bool]$sv.windows_auth_only; Volltext = [bool]$sv.fulltext_installed
				'TCP-Port' = $port; 'Uptime h' = $sv.uptime_hours; 'SQL-Start' = $sv.sqlserver_start_time
				'SQL CPU avg %' = $ld.sql_cpu_avg_pct; 'SQL CPU p95 %' = $ld.sql_cpu_p95_pct; 'SQL CPU max %' = $ld.sql_cpu_max_pct; 'Gesamt CPU p95 %' = $ld.total_cpu_p95_pct; 'CPU Historie min' = $ld.cpu_samples_minutes
				'Page Life Expectancy s' = $ld.page_life_expectancy_sec; 'Memory aktuell MB' = $ld.total_server_memory_mb; 'Memory Ziel MB' = $ld.target_server_memory_mb
				'Batch Requests/s (Messung)' = $ld.batch_requests_per_sec_sampled; 'Batch Requests/s (avg)' = $ld.batch_requests_per_sec_avg_since_start
				'Transaktionen/s (Messung)' = $ld.transactions_per_sec_sampled; 'Transaktionen/s (avg)' = $ld.transactions_per_sec_avg_since_start; Verbindungen = $ld.user_connections
				'Anzahl DBs (ohne tempdb)' = $userDbs.Count; 'DB-Daten GB' = $dataGb
				'Agent-Jobs' = (Get-Prop $i.agent_jobs_summary 'total'); 'Agent-Jobs aktiv' = (Get-Prop $i.agent_jobs_summary 'enabled'); 'Linked Server' = @($i.linked_servers).Count
				'DB-Mail-Profile' = (Get-Prop $i.components 'database_mail_profiles'); Credentials = @($i.components.credentials).Count; Proxies = @($i.components.proxies).Count
				'Availability Group' = $agNames; 'AG-Rolle (lokal)' = $agLocal; 'AG-Partnerknoten' = $agPartners
				'xp_cmdshell' = (Get-Prop $cfg 'xp_cmdshell'); 'CLR aktiv' = (Get-Prop $cfg 'clr enabled')
			})

		foreach ($d in $dbs)
		{
			$b = $d.backup
			Add-Row $S_db ([ordered]@{
					Server = $hn; Anwendung = $app; Obergruppe = $grp; Instanz = $i.instance_name; Datenbank = $d.name; 'Availability Group' = $agDbMap[[string]$d.name]; Status = $d.state
					Compat = $d.compatibility_level; Recovery = $d.recovery_model; Collation = $d.collation; Owner = $d.owner
					'Daten GB' = (ToGb $d.data_size_mb 2); 'Daten belegt GB' = (ToGb $d.data_used_mb 2); 'Log GB' = (ToGb $d.log_size_mb 2); 'Log belegt %' = $d.log_used_pct
					Dateien = $d.file_count; Tabellen = $d.table_count; 'Wachstum MB/Monat' = $d.growth_mb_per_month; 'Full-Groessen MB/Monat' = $d.full_backup_size_history_mb
					TDE = $d.is_encrypted_tde; 'Always Encrypted Keys' = $d.always_encrypted_keys; 'CLR-Assemblies' = $d.clr_assemblies; Containment = $d.containment; 'Read-Only' = $d.is_read_only; Snapshot = $d.is_snapshot
					'Full letzte' = $b.full.last_finish; 'Full Anzahl' = $b.full.count_in_window; 'Full avg min' = (ToMin $b.full.avg_duration_sec); 'Full max min' = (ToMin $b.full.max_duration_sec)
					'Full avg GB' = (ToGb $b.full.avg_size_mb 2); 'Full komprimiert GB' = (ToGb $b.full.avg_compressed_mb 2); 'Full Ziel' = $b.full.last_target
					'Diff letzte' = $b.diff.last_finish; 'Diff avg min' = (ToMin $b.diff.avg_duration_sec)
					'Log letzte' = $b.log.last_finish; 'Log Anzahl' = $b.log.count_in_window; 'Log avg min' = (ToMin $b.log.avg_duration_sec)
					Publisher = $d.is_published; Subscriber = $d.is_subscribed; 'Merge-Publisher' = $d.is_merge_published; Distributor = $d.is_distributor; CDC = $d.cdc_enabled; 'Service Broker' = $d.service_broker
				})
		}
		foreach ($f in (AsList $i.files))
		{
			Add-Row $S_files ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Datenbank = $f.database; 'Logischer Name' = $f.logical_name; Typ = $f.type; Pfad = $f.path; Laufwerk = $f.volume; 'Groesse MB' = $f.size_mb; 'Growth' = $f.growth; 'Growth in %' = $f.percent_growth; 'Max Size' = $f.max_size })
		}

		# dependencies
		foreach ($ls in (AsList $i.linked_servers))
		{
			Add-Row $S_dep ([ordered]@{ Typ = 'Linked Server'; Server = $hn; Instanz = $i.instance_name; Quelle = $null; Ziel = $ls.name; Detail = "$($ls.provider) / $($ls.data_source)"; Info = "Login-Mappings: $($ls.login_mappings), feste Remote-Logins: $($ls.fixed_remote_logins)" })
		}
		foreach ($x in (AsList $i.cross_database_refs))
		{
			if ($x.ref_server) { Add-Row $S_dep ([ordered]@{ Typ = 'Cross-Server-Referenz'; Server = $hn; Instanz = $i.instance_name; Quelle = $x.database; Ziel = $x.ref_server; Detail = $x.ref_database; Info = "Referenzen: $($x.ref_count)" }) }
			else { Add-Row $S_dep ([ordered]@{ Typ = 'Cross-DB-Referenz'; Server = $hn; Instanz = $i.instance_name; Quelle = $x.database; Ziel = $x.ref_database; Detail = $null; Info = "Referenzen: $($x.ref_count)" }) }
		}
		foreach ($d in $dbs)
		{
			if ($d.is_published -or $d.is_merge_published) { Add-Row $S_dep ([ordered]@{ Typ = 'Replikation (Publisher)'; Server = $hn; Instanz = $i.instance_name; Quelle = $d.name; Ziel = $null; Detail = $null; Info = $null }) }
			if ($d.is_subscribed) { Add-Row $S_dep ([ordered]@{ Typ = 'Replikation (Subscriber)'; Server = $hn; Instanz = $i.instance_name; Quelle = $d.name; Ziel = $null; Detail = $null; Info = $null }) }
			if ($d.is_distributor) { Add-Row $S_dep ([ordered]@{ Typ = 'Replikation (Distributor)'; Server = $hn; Instanz = $i.instance_name; Quelle = $d.name; Ziel = $null; Detail = $null; Info = $null }) }
		}
		if ($i.agent_jobs)
		{
			$anyJobs = $true
			foreach ($j in (AsList $i.agent_jobs))
			{
				Add-Row $S_jobs ([ordered]@{
						Server = $hn; Instanz = $i.instance_name; Job = $j.name; Aktiv = $j.enabled; Owner = $j.owner; Kategorie = $j.category; Schritte = $j.step_count; Subsysteme = $j.subsystems
						'Proxy genutzt' = $j.uses_proxy; Zeitplan = $j.schedules; 'Letzter Lauf' = $j.last_run; 'Dauer avg s' = $j.avg_duration_sec; 'Fehlschlaege (Historie)' = $j.failures_in_history; 'Externe Referenzen' = $j.external_refs
					})
				if ($j.external_refs) { Add-Row $S_dep ([ordered]@{ Typ = 'Agent-Job extern'; Server = $hn; Instanz = $i.instance_name; Quelle = $j.name; Ziel = $j.external_refs; Detail = $j.subsystems; Info = $null }) }
			}
		}

		# HA / DR
		$ha = $i.high_availability
		foreach ($r in (AsList $ha.availability_replicas)) { Add-Row $S_ha ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'AG-Replikat'; Name = $r.ag_name; Detail = $r.replica_server_name; Info = "$($r.role_desc) / $($r.availability_mode_desc) / $($r.failover_mode_desc) / Secondary: $($r.secondary_role_allow_connections_desc)" }) }
		foreach ($r in (AsList $ha.listeners)) { Add-Row $S_ha ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'AG-Listener'; Name = $r.ag_name; Detail = $r.dns_name; Info = "Port $($r.port)" }) }
		foreach ($r in (AsList $ha.availability_databases)) { Add-Row $S_ha ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'AG-Datenbank'; Name = $r.ag_name; Detail = $r.database_name; Info = $null }) }
		foreach ($r in (AsList $ha.cluster_nodes)) { Add-Row $S_ha ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'Cluster-Knoten'; Name = $r.node_name; Detail = $r.status_description; Info = "Aktueller Besitzer: $($r.is_current_owner)" }) }
		foreach ($r in (AsList $ha.log_shipping_primary)) { Add-Row $S_ha ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'Log Shipping (Primary)'; Name = $r.primary_database; Detail = $null; Info = $null }) }
		foreach ($r in (AsList $ha.log_shipping_secondary)) { Add-Row $S_ha ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'Log Shipping (Secondary)'; Name = $r.secondary_database; Detail = $null; Info = $null }) }
		foreach ($r in (AsList $ha.mirroring)) { Add-Row $S_ha ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'Database Mirroring'; Name = $r.database_name; Detail = $r.mirroring_partner_name; Info = "$($r.mirroring_role_desc) / $($r.mirroring_state_desc)" }) }

		# components
		foreach ($r in (AsList $i.components.credentials)) { Add-Row $S_comp ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'Credential'; Name = $r.name; Detail = $r.credential_identity; Konto = $null }) }
		foreach ($r in (AsList $i.components.proxies)) { Add-Row $S_comp ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'Agent-Proxy'; Name = $r.name; Detail = $r.credential_identity; Konto = $null }) }
		foreach ($r in (AsList $i.components.certificates)) { Add-Row $S_comp ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'Zertifikat (master)'; Name = $r.name; Detail = "gueltig bis $($r.expiry_date)"; Konto = $r.subject }) }
		foreach ($r in (AsList $i.components.tde_databases)) { Add-Row $S_comp ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'TDE'; Name = $r.database; Detail = "$($r.algorithm) $($r.key_length) / State $($r.encryption_state)"; Konto = $null }) }
		foreach ($r in (AsList $i.endpoints)) { Add-Row $S_comp ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Typ = 'Endpoint'; Name = $r.name; Detail = "$($r.type_desc) / $($r.state_desc)"; Konto = $(if ($r.port) { "Port $($r.port)" }) }) }
	}
}

# summary per parent group
$S_sum = New-Object System.Collections.ArrayList
$serverRows = @($S_server)
$totalServers = $serverRows.Count
foreach ($g in @($serverRows | Group-Object { $_.Obergruppe } | Sort-Object Name))
{
	$names = @($g.Group | ForEach-Object { $_.Server })
	$mySrv = $g.Group
	$myInst = @($S_inst | Where-Object { $_.Server -in $names })
	$myDb = @($S_db | Where-Object { $_.Server -in $names -and $_.Datenbank -ne 'tempdb' })
	$myDisk = @($S_disks | Where-Object { $_.Server -in $names })
	Add-Row $S_sum ([ordered]@{
			Obergruppe = $g.Name; Server = $names.Count; 'Anteil Server %' = $(if ($totalServers) { [math]::Round(100.0 * $names.Count / $totalServers, 1) } else { $null })
			Instanzen = $myInst.Count; 'Datenbanken (ohne tempdb)' = $myDb.Count
			'Kerne gesamt' = ($mySrv | Measure-Object -Property Kerne -Sum).Sum; 'Logische CPUs' = ($mySrv | Measure-Object -Property 'Logische CPUs' -Sum).Sum
			'RAM GB' = [math]::Round(($mySrv | Measure-Object -Property 'RAM GB' -Sum).Sum, 0)
			'Platten GB (Kapazitaet)' = [math]::Round(($myDisk | Measure-Object -Property 'Groesse GB' -Sum).Sum, 0)
			'Platten GB (belegt)' = [math]::Round(($myDisk | Measure-Object -Property 'Belegt GB' -Sum).Sum, 0)
			'DB-Daten GB' = [math]::Round(($myDb | Measure-Object -Property 'Daten GB' -Sum).Sum, 1)
			'DB-Daten belegt GB' = [math]::Round(($myDb | Measure-Object -Property 'Daten belegt GB' -Sum).Sum, 1)
			'Wachstum GB/Monat' = [math]::Round((($myDb | Measure-Object -Property 'Wachstum MB/Monat' -Sum).Sum) / 1024, 1)
			'Virtuelle Server' = @($mySrv | Where-Object { $_.Virtuell -eq $true }).Count
			'Server ohne Anwendung' = @($mySrv | Where-Object { $_.Anwendung -eq 'NICHT ZUGEORDNET' }).Count
		})
}
if ($S_sum.Count -gt 1)
{
	$all = @($S_sum)
	Add-Row $S_sum ([ordered]@{
			Obergruppe = 'GESAMT'; Server = $totalServers; 'Anteil Server %' = 100
			Instanzen = ($all | Measure-Object -Property Instanzen -Sum).Sum; 'Datenbanken (ohne tempdb)' = ($all | Measure-Object -Property 'Datenbanken (ohne tempdb)' -Sum).Sum
			'Kerne gesamt' = ($all | Measure-Object -Property 'Kerne gesamt' -Sum).Sum; 'Logische CPUs' = ($all | Measure-Object -Property 'Logische CPUs' -Sum).Sum
			'RAM GB' = ($all | Measure-Object -Property 'RAM GB' -Sum).Sum
			'Platten GB (Kapazitaet)' = ($all | Measure-Object -Property 'Platten GB (Kapazitaet)' -Sum).Sum; 'Platten GB (belegt)' = ($all | Measure-Object -Property 'Platten GB (belegt)' -Sum).Sum
			'DB-Daten GB' = ($all | Measure-Object -Property 'DB-Daten GB' -Sum).Sum; 'DB-Daten belegt GB' = ($all | Measure-Object -Property 'DB-Daten belegt GB' -Sum).Sum
			'Wachstum GB/Monat' = [math]::Round(($all | Measure-Object -Property 'Wachstum GB/Monat' -Sum).Sum, 1)
			'Virtuelle Server' = ($all | Measure-Object -Property 'Virtuelle Server' -Sum).Sum; 'Server ohne Anwendung' = ($all | Measure-Object -Property 'Server ohne Anwendung' -Sum).Sum
		})
}
#endregion

#region anonymize
$anonMap = @{ }
$anonRegex = $null
function Invoke-Anonymize
{
	param ($Sheets)
	$ids = New-Object System.Collections.ArrayList
	$n = 0
	foreach ($s in $scans)
	{
		$n++
		$id = 'SRV{0:D3}' -f $n
		$anonMap[$s.host.computer_name.ToUpperInvariant()] = $id
		[void]$ids.Add([pscustomobject]@{ Original = $s.host.computer_name; Kennung = $id; Typ = 'Server' })
		foreach ($i in (AsList $s.instances))
		{
			if ($i.instance_name -and $i.instance_name -ne 'MSSQLSERVER') { if (-not $anonMap.ContainsKey($i.instance_name.ToUpperInvariant())) { $anonMap[$i.instance_name.ToUpperInvariant()] = "$id-INST$($anonMap.Count)"; [void]$ids.Add([pscustomobject]@{ Original = $i.instance_name; Kennung = $anonMap[$i.instance_name.ToUpperInvariant()]; Typ = 'Instanz' }) } }
		}
		$dom = [string]$s.host.domain
		if ($dom -and $dom -ne 'WORKGROUP' -and -not $anonMap.ContainsKey($dom.ToUpperInvariant())) { $anonMap[$dom.ToUpperInvariant()] = 'DOMAIN'; [void]$ids.Add([pscustomobject]@{ Original = $dom; Kennung = 'DOMAIN'; Typ = 'Domaene' }) }
	}
	$alts = @($anonMap.Keys | Sort-Object { $_.Length } -Descending | ForEach-Object { [regex]::Escape($_) })
	$rx = New-Object System.Text.RegularExpressions.Regex(('(?<![A-Za-z0-9_])(' + ($alts -join '|') + ')(?![A-Za-z0-9_])'), 'IgnoreCase')
	$ev = [System.Text.RegularExpressions.MatchEvaluator]{ param ($m) $anonMap[$m.Value.ToUpperInvariant()] }
	foreach ($name in (AsList $Sheets.Keys))
	{
		foreach ($row in @($Sheets[$name]))
		{
			foreach ($p in (AsList $row.PSObject.Properties))
			{
				if ($p.Value -is [string] -and $p.Value.Length -gt 0)
				{
					if ($p.Name -in @('FQDN', 'IPs')) { $p.Value = $null; continue }
					$p.Value = $rx.Replace($p.Value, $ev)
				}
			}
		}
	}
	return $ids
}
#endregion

#region write provider workbook
$provider = [ordered]@{ }
$provider['Zusammenfassung'] = @($S_sum)
$provider['Server'] = @($S_server)
$provider['Platten'] = @($S_disks)
$provider['Datentraeger'] = @($S_hw)
$provider['SQL-Instanzen'] = @($S_inst)
$provider['Datenbanken'] = @($S_db)
$provider['Dateien'] = @($S_files)
$provider['Abhaengigkeiten'] = @($S_dep)
$provider['HA-DR'] = @($S_ha)
$provider['Komponenten'] = @($S_comp)
if ($anyJobs) { $provider['Agent-Jobs'] = @($S_jobs) }

$anonFile = $null
if ($Anonymize)
{
	$ids = Invoke-Anonymize -Sheets $provider
	$anonFile = Join-Path $OutputPath "Anonymize_Zuordnung_$stamp.csv"
	$ids | Export-Csv -LiteralPath $anonFile -Delimiter ';' -NoTypeInformation -Encoding UTF8
}
$providerFile = Join-Path $OutputPath "Provider_Inventory_$stamp.xlsx"
Write-XlsxFile -Path $providerFile -Sheets $provider
#endregion

#region security workbook
$securityFile = $null
if ($IncludeSecurity)
{
	$S_login = New-Object System.Collections.ArrayList
	$S_srvperm = New-Object System.Collections.ArrayList
	$S_dbuser = New-Object System.Collections.ArrayList
	$S_dbrole = New-Object System.Collections.ArrayList
	$S_dbmem = New-Object System.Collections.ArrayList
	$S_dbperm = New-Object System.Collections.ArrayList
	$S_secsum = New-Object System.Collections.ArrayList
	foreach ($s in $scans)
	{
		$hn = $s.host.computer_name
		$app = Get-MapValue $s 'Anwendung' 'application'
		$grp = Get-MapValue $s 'Obergruppe' 'parent_group'
		foreach ($i in (AsList $s.instances))
		{
			if (-not $i.connected) { continue }
			if ($null -eq $i.logins -and $s.options.skip_security) { Add-Report 'HINWEIS' $hn 'Scan wurde mit -SkipSecurity erstellt, keine Security-Daten.'; continue }
			$excluded = @{ }
			$kept = 0
			foreach ($l in (AsList $i.logins))
			{
				$isEx = (-not $IncludeSystemAndSysadmin) -and ($l.is_system -or $l.is_sysadmin)
				if ($isEx) { $excluded[[string]$l.name] = $true; continue }
				$kept++
				Add-Row $S_login ([ordered]@{
						Server = $hn; Anwendung = $app; Obergruppe = $grp; Instanz = $i.instance_name; Login = $l.name; Typ = $l.type; Deaktiviert = $l.disabled; 'Default-DB' = $l.default_database
						'Server-Rollen' = $l.server_roles; sysadmin = $l.is_sysadmin; 'Policy geprueft' = $l.policy_checked; 'Ablauf geprueft' = $l.expiration_checked; Erstellt = $l.created; Geaendert = $l.modified
					})
			}
			foreach ($p in (AsList $i.server_permissions)) { if (-not $excluded.ContainsKey([string]$p.grantee)) { Add-Row $S_srvperm ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Grantee = $p.grantee; Recht = $p.permission_name; Status = $p.state_desc }) } }
			$orph = 0
			foreach ($u in (AsList $i.db_users))
			{
				if ($u.login_name -and $excluded.ContainsKey([string]$u.login_name)) { continue }
				if ($excluded.ContainsKey([string]$u.name)) { continue }
				if ($u.orphaned) { $orph++ }
				Add-Row $S_dbuser ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Datenbank = $u.database; 'DB-User' = $u.name; Typ = $u.type; Authentifizierung = $u.authentication; Login = $u.login_name; 'Default-Schema' = $u.default_schema; Verwaist = $u.orphaned; Erstellt = $u.create_date })
			}
			foreach ($r in (AsList $i.db_roles)) { Add-Row $S_dbrole ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Datenbank = $r.database; Rolle = $r.name; Typ = $r.type; Owner = $r.owner }) }
			foreach ($m in (AsList $i.db_role_members))
			{
				if ($excluded.ContainsKey([string]$m.member)) { continue }
				Add-Row $S_dbmem ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Datenbank = $m.database; Rolle = $m.role; 'Feste Rolle' = $m.role_is_fixed; Mitglied = $m.member; 'Mitglied-Typ' = $m.member_type })
			}
			foreach ($p in (AsList $i.db_permissions))
			{
				if ($excluded.ContainsKey([string]$p.grantee)) { continue }
				Add-Row $S_dbperm ([ordered]@{ Server = $hn; Instanz = $i.instance_name; Datenbank = $p.database; Grantee = $p.grantee; 'Grantee-Typ' = $p.grantee_type; Klasse = $p.class; Recht = $p.permission; Status = $p.state; Objekt = $p.target })
			}
			$lg = @($i.logins | Where-Object { -not $excluded.ContainsKey([string]$_.name) })
			Add-Row $S_secsum ([ordered]@{
					Server = $hn; Obergruppe = $grp; Instanz = $i.instance_name; 'Logins (gefiltert)' = $kept
					'Windows-Logins' = @($lg | Where-Object { $_.type -eq 'WINDOWS_LOGIN' }).Count; 'Windows-Gruppen (AD)' = @($lg | Where-Object { $_.type -eq 'WINDOWS_GROUP' }).Count
					'SQL-Logins' = @($lg | Where-Object { $_.type -eq 'SQL_LOGIN' }).Count; 'Zertifikat/Key-Logins' = @($lg | Where-Object { $_.type -in @('CERTIFICATE_MAPPED_LOGIN', 'ASYMMETRIC_KEY_MAPPED_LOGIN') }).Count
					'Ausgefilterte (System/sysadmin)' = $excluded.Count; 'DB-User' = @($S_dbuser | Where-Object { $_.Server -eq $hn -and $_.Instanz -eq $i.instance_name }).Count; 'Verwaiste DB-User' = $orph
					'Eigene DB-Rollen' = @($i.db_roles).Count
				})
		}
	}
	$sec = [ordered]@{ }
	$sec['Zusammenfassung'] = @($S_secsum)
	$sec['Logins'] = @($S_login)
	$sec['Server-Rechte'] = @($S_srvperm)
	$sec['DB-User'] = @($S_dbuser)
	$sec['DB-Rollen'] = @($S_dbrole)
	$sec['Rollenmitglieder'] = @($S_dbmem)
	$sec['DB-Berechtigungen'] = @($S_dbperm)
	$securityFile = Join-Path $OutputPath "Security_Inventory_$stamp.xlsx"
	Write-XlsxFile -Path $securityFile -Sheets $sec
}
#endregion

#region report
$reportFile = Join-Path $OutputPath "Pruefbericht_$stamp.txt"
$lines = New-Object System.Collections.ArrayList
[void]$lines.Add("Pruefbericht Provider-Inventar, $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
[void]$lines.Add("Scans: $($scans.Count) Server, Mapping-Zeilen: $($mappingRows.Count)")
[void]$lines.Add('')
foreach ($grpL in @($report | Group-Object Stufe | Sort-Object Name))
{
	[void]$lines.Add("== $($grpL.Name) ($($grpL.Count)) ==")
	foreach ($r in $grpL.Group) { [void]$lines.Add("  $($r.Server): $($r.Meldung)") }
	[void]$lines.Add('')
}
if ($report.Count -eq 0) { [void]$lines.Add('Keine Auffaelligkeiten.') }
[System.IO.File]::WriteAllText($reportFile, ($lines -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))

Write-Host "Servers: $($scans.Count)   Report entries: $($report.Count)"
Write-Host "Provider workbook : $providerFile"
if ($securityFile) { Write-Host "Security workbook : $securityFile (NOT for the provider)" }
if ($anonFile) { Write-Host "Anonymize key     : $anonFile (NOT for the provider)" }
Write-Host "Check report      : $reportFile"
if ($report.Count -gt 0) { Write-Warning "$($report.Count) finding(s), see check report." }
#endregion
