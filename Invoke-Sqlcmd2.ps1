function Invoke-Sqlcmd2 {
<#
.SYNOPSIS
Simpler alternative to Microsoft's Invoke-Sqlcmd

.DESCRIPTION
Similar syntax to Microsoft's Invoke-Sqlcmd but has no dependencies and always returns an array, even for a single row of data.

.EXAMPLE
Invoke-Sqlcmd2 -ServerInstance localhost -Database SystemFrontierV2 -TrustServerCertificate -Query "select top 10 * from [dbo].[Computer]"

.LINK
https://github.com/systemfrontier/Invoke-Sqlcmd2

.NOTES
==============================================
Version:	2.1
Author:		Jay Adams, Noxigen LLC
Created:	2024-04-21
Copyright:	Noxigen LLC. All rights reserved.

Create secure GUIs for PowerShell with System Frontier.
https://systemfrontier.com/
==============================================
History:	2024-07-17	Jay Adams, Noxigen LLC
			Fix: Single result with a single row not being treated as an array of DataRow
			Feat: Multiple results sets are returned as arrays of DataRow for more uniformity across result types
			
			2024-07-30	Jay Adams, Noxigen LLC
			Fix: Single result set returns ArrayList of rows now
#>
	param (
		[Parameter(Mandatory=$true)]
		[string]
		$ServerInstance,
		
		[Parameter(Mandatory=$true)]
		[string]
		[ValidateNotNullOrEmpty()]
		$Database,
		
		[Parameter(Mandatory=$true, ParameterSetName='QuerySet1')]
		[string]
		[ValidateNotNullOrEmpty()]
		$Query,

		[Parameter(Mandatory=$true, ParameterSetName='QuerySet2')]
		[string]
		[ValidateNotNullOrEmpty()]
		$InputFile,

		[Parameter()]
		[ValidateSet("SingleResultSet","MultipleResultSets","Scalar","NonQuery")]
		[string]
		$QueryType = "none",

		[Parameter()]
		[switch]
		$Unencrypted = $false,

		[Parameter()]
		[switch]
		$TrustServerCertificate = $false,

		[Parameter()]
		[ValidateRange(1, 65535)]
		[int]
		$QueryTimeout = 120,

		[Parameter()]
		[ValidateRange(1, 65534)]
		[int]
		$ConnectionTimeout = 30,

		[Parameter()]
		[string]
		$ApplicationName,

		[Parameter()]
		[string]
		$ConnectionString
	)

	if ($QueryType -eq "none" -and $Query -match ";") {
		$QueryType = "MultipleResultSets"
	}

	if ($InputFile) {
		if ((Test-Path -Path $InputFile -Type Leaf) -eq $true) {
			$Query = Get-Content -Path $InputFile -Raw

			if ([string]::IsNullOrWhiteSpace($Query)) {
				throw "Input file is invalid"
			}
		} else {
			throw "Input file not found"
		}
	}

	if (![string]::IsNullOrWhiteSpace($ConnectionString)) {
		$_connectionString = $ConnectionString
	} else {
		$_connectionString = `
		"Data Source=$ServerInstance;" + 
		"Initial Catalog=$Database;" + 
		"Integrated Security=True;" + 
		"Encrypt=$(-not $Unecrypted);" + 
		"TrustServerCertificate=$TrustServerCertificate;" + 
		"Connection Timeout=$ConnectionTimeout;"

		if (![string]::IsNullOrWhitespace($ApplicationName)) {
			$_connectionString += "Application Name=$ApplicationName;"
		}
	}

	$connection = [System.Data.SqlClient.SqlConnection]::new($_connectionString)
	
	$command = [System.Data.SqlClient.SqlCommand]::new($Query, $connection)
	$command.CommandTimeout = $QueryTimeout
	
	try {
		$connection.Open()

		if ($QueryType -eq "NonQuery") {
			$command.ExecuteNonQuery()
		} else {
			if ($QueryType -eq "Scalar") {
				return $command.ExecuteScalar()
			} else {
				$adapter = [System.Data.SqlClient.SqlDataAdapter]::new($command)
				$dataset = [System.Data.DataSet]::new()
				$adapter.Fill($dataset) | Out-Null
				$connection.Close();

				if ($null -ne $dataset) {
					if ($QueryType -eq "MultipleResultSets") {
						$rs = New-Object System.Collections.ArrayList

						foreach ($table in $dataset.Tables)
						{
							[void]$rs.Add($table.Rows)
						}

						return $rs
					} else {
						# Single result set
						$results = New-Object System.Collections.ArrayList
						[void]$results.Add($dataset.Tables[0].Rows)
						return $results
					}
				}
			}
		}
	} finally {
		if ($connection.State -eq [System.Data.ConnectionState]::Open) {
			$connection.Close()
		}
	}
}
