#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
# Run with PowerShell 7: Invoke-Pester .\tests\Merge.Tests.ps1
# End-to-end test of Merge-sqmProviderInventory.ps1 against a synthetic fixture scan.

BeforeAll {
	$script:Root = Split-Path -Parent $PSScriptRoot
	$script:Merge = Join-Path $script:Root 'Merge-sqmProviderInventory.ps1'
	$script:Fixtures = Join-Path $PSScriptRoot 'fixtures'
	Add-Type -AssemblyName System.IO.Compression.FileSystem

	function script:Get-XlsxSheetNames
	{
		param ([string]$Path)
		$zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
		try
		{
			$e = $zip.GetEntry('xl/workbook.xml')
			$sr = New-Object System.IO.StreamReader($e.Open())
			$xml = [xml]$sr.ReadToEnd()
			$sr.Dispose()
			return @($xml.workbook.sheets.sheet | ForEach-Object { $_.name })
		}
		finally { $zip.Dispose() }
	}
	function script:Get-XlsxAllText
	{
		param ([string]$Path)
		$zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
		try
		{
			$sb = New-Object System.Text.StringBuilder
			foreach ($e in $zip.Entries)
			{
				$sr = New-Object System.IO.StreamReader($e.Open())
				[void]$sb.Append($sr.ReadToEnd())
				$sr.Dispose()
			}
			return $sb.ToString()
		}
		finally { $zip.Dispose() }
	}
}

Describe 'Merge-sqmProviderInventory' {
	BeforeEach {
		$script:Out = Join-Path ([System.IO.Path]::GetTempPath()) ("sqmpi_" + [guid]::NewGuid().ToString('N'))
		New-Item -ItemType Directory -Path $script:Out | Out-Null
	}
	AfterEach { Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue }

	It 'writes the provider workbook without any security sheet or login name' {
		& $script:Merge -InputPath $script:Fixtures -OutputPath $script:Out 6>$null 3>$null
		$f = Get-ChildItem $script:Out -Filter 'Provider_Inventory_*.xlsx'
		$f | Should -Not -BeNullOrEmpty
		$names = Get-XlsxSheetNames $f.FullName
		$names | Should -Contain 'Datenbanken'
		$names | Should -Contain 'Zusammenfassung'
		($names -join ',') | Should -Not -Match 'Login|Security|DB-User'
		$text = Get-XlsxAllText $f.FullName
		$text | Should -Not -Match 'appsvc'
		$text | Should -Not -Match 'dba-group'
	}

	It 'does not write a security workbook unless asked' {
		& $script:Merge -InputPath $script:Fixtures -OutputPath $script:Out 6>$null 3>$null
		Get-ChildItem $script:Out -Filter 'Security_Inventory_*.xlsx' | Should -BeNullOrEmpty
	}

	It 'filters sysadmin and system logins in the security workbook by default' {
		& $script:Merge -InputPath $script:Fixtures -OutputPath $script:Out -IncludeSecurity 6>$null 3>$null
		$f = Get-ChildItem $script:Out -Filter 'Security_Inventory_*.xlsx'
		$f | Should -Not -BeNullOrEmpty
		$text = Get-XlsxAllText $f.FullName
		$text | Should -Match 'appsvc'
		$text | Should -Not -Match 'dba-group'
		$text | Should -Not -Match '>sa<'
	}

	It 'keeps sysadmin logins with -IncludeSystemAndSysadmin' {
		& $script:Merge -InputPath $script:Fixtures -OutputPath $script:Out -IncludeSecurity -IncludeSystemAndSysadmin 6>$null 3>$null
		$f = Get-ChildItem $script:Out -Filter 'Security_Inventory_*.xlsx'
		(Get-XlsxAllText $f.FullName) | Should -Match 'dba-group'
	}

	It 'anonymizes host name and domain and writes the key file' {
		& $script:Merge -InputPath $script:Fixtures -OutputPath $script:Out -Anonymize 6>$null 3>$null
		$f = Get-ChildItem $script:Out -Filter 'Provider_Inventory_*.xlsx'
		$text = Get-XlsxAllText $f.FullName
		$text | Should -Not -Match 'FIXTURE01'
		$text | Should -Not -Match 'example\.test'
		$text | Should -Match 'SRV001'
		Get-ChildItem $script:Out -Filter 'Anonymize_Zuordnung_*.csv' | Should -Not -BeNullOrEmpty
	}

	It 'reports servers missing in the mapping and mapping rows without scan' {
		$map = Join-Path $script:Out 'Mapping.csv'
		"Server;Anwendung;Obergruppe`r`nOTHERSRV;X;Y" | Set-Content -LiteralPath $map -Encoding UTF8
		& $script:Merge -InputPath $script:Fixtures -MappingPath $map -OutputPath $script:Out 6>$null 3>$null
		$rep = Get-Content -LiteralPath (Get-ChildItem $script:Out -Filter 'Pruefbericht_*.txt').FullName -Raw
		$rep | Should -Match 'OTHERSRV: Im Mapping, aber kein Scan'
		$rep | Should -Match 'FIXTURE01: Server fehlt im Mapping'
	}

	It 'the mapping CSV wins over values set at scan time' {
		$map = Join-Path $script:Out 'Mapping.csv'
		"Server;Anwendung;Obergruppe`r`nFIXTURE01;CsvApp;CsvGroup" | Set-Content -LiteralPath $map -Encoding UTF8
		& $script:Merge -InputPath $script:Fixtures -MappingPath $map -OutputPath $script:Out 6>$null 3>$null
		$text = Get-XlsxAllText (Get-ChildItem $script:Out -Filter 'Provider_Inventory_*.xlsx').FullName
		$text | Should -Match 'CsvApp'
		$text | Should -Not -Match 'AppA'
	}

	It 'records the AG partner node and reports it when the partner was not scanned' {
		& $script:Merge -InputPath $script:Fixtures -OutputPath $script:Out 6>$null 3>$null
		$text = Get-XlsxAllText (Get-ChildItem $script:Out -Filter 'Provider_Inventory_*.xlsx').FullName
		$text | Should -Match 'FIXTURE02'
		$text | Should -Match 'PRIMARY'
		$text | Should -Match 'AG1LSN'
		$rep = Get-Content -LiteralPath (Get-ChildItem $script:Out -Filter 'Pruefbericht_*.txt').FullName -Raw
		$rep | Should -Match 'FIXTURE02: AG-Partnerknoten von FIXTURE01'
	}

	It 'reads the fixed TCP port and ignores an empty dynamic port' {
		& $script:Merge -InputPath $script:Fixtures -OutputPath $script:Out 6>$null 3>$null
		$text = Get-XlsxAllText (Get-ChildItem $script:Out -Filter 'Provider_Inventory_*.xlsx').FullName
		$text | Should -Match '>1433<'
		$text | Should -Not -Match 'dynamisch'
	}

	It 'creates and extends a mapping template without losing existing rows' {
		$map = Join-Path $script:Out 'Mapping.csv'
		"Server;Anwendung;Obergruppe`r`nOTHERSRV;X;Y" | Set-Content -LiteralPath $map -Encoding UTF8
		& $script:Merge -InputPath $script:Fixtures -MappingPath $map -NewMappingTemplate 6>$null 3>$null
		$rows = Import-Csv -LiteralPath $map -Delimiter ';' -Encoding UTF8
		@($rows).Count | Should -Be 2
		($rows | Where-Object Server -eq 'OTHERSRV').Anwendung | Should -Be 'X'
		($rows | Where-Object Server -eq 'FIXTURE01').Obergruppe | Should -Be 'GroupX'
	}
}
