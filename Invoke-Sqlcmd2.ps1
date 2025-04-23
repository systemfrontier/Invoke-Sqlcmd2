function Invoke-Sqlcmd2 {
    <#
    .SYNOPSIS
    Simpler alternative to Microsoft's Invoke-Sqlcmd
    
    .DESCRIPTION
    Similar syntax to Microsoft's Invoke-Sqlcmd but has no dependencies and always returns an array, even for a single row of data.
    
    .EXAMPLE
    Invoke-Sqlcmd2 -Credential $v2SqlCred -ServerInstance localhost -Database SystemFrontierV2 -TrustServerCertificate -Query "select top 10 * from [dbo].[Computer]"
    
    .LINK
    https://github.com/systemfrontier/Invoke-Sqlcmd2
    
    .NOTES
    ==============================================
    Version:	2.3
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
                2024-09-18	Jay Adams, Noxigen LLC
                Feat: Added support for SQL authentication and parameterized queries
                2025-04-23	Jay Adams, Noxigen LLC
                Feat: Support showing PRINT messages in console output
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
        $ConnectionString,

        [Parameter()]
        [PSCredential]
        $Credential,

        [Parameter()]
        [System.Collections.Specialized.OrderedDictionary]
        $Parameters,

        [Parameter()]
        [switch]
        $ShowMessages = $false
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
        "Encrypt=$(-not $Unecrypted);" + 
        "Connection Timeout=$ConnectionTimeout;"

        if (![string]::IsNullOrWhitespace($ApplicationName)) {
            $_connectionString += "Application Name=$ApplicationName;"
        }

        if ($null -eq $Credential) {
            $_connectionString += "Integrated Security=True;"
        } else {
            $_connectionString += "User Id=$($Credential.UserName);Password=$($Credential.GetNetworkCredential().Password);"
        }

        $_connectionString += "TrustServerCertificate=$TrustServerCertificate;"
    }

    $connection = [System.Data.SqlClient.SqlConnection]::new($_connectionString)
    
    $command = [System.Data.SqlClient.SqlCommand]::new()
    $command.Connection = $connection
    $command.CommandTimeout = $QueryTimeout

    if ($null -ne $Parameters -and $Parameters.Count -gt 0) {
        
        foreach ($parameter in $Parameters.GetEnumerator()) {
            if ($null -ne $parameter.Value) {
                $command.Parameters.AddWithValue("@$($parameter.Name)", $parameter.Value) | Out-Null
            }
        }
    }

    $command.CommandText = $Query

    try {
        $connection.Open()

        Register-ObjectEvent -InputObject $connection -MessageData $ShowMessages -EventName InfoMessage `
        -Action { if ($event.MessageData -eq $true) { Write-Host "$($event.SourceEventArgs)" } } -SupportEvent

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
