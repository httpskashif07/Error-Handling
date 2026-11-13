function Get-ProcessInfo {
<#
.SYNOPSIS
    Retrieves information about specified processes.
.DESCRIPTION
    This function takes a list of process names and displays information about each one.
.PARAMETER Processes
    An array of process names to check.
.EXAMPLE
    Get-ProcessInfo -Processes "lsass","spoolsv","dns","w3wp"
#>

 [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Processes
    )

    $Results = @()
    $Found   = 0
    $Missing = 0

    foreach ($Process in $Processes)
    {
        Write-Verbose "Checking process: $Process"
        $Result = Get-Process $Process -ErrorAction SilentlyContinue

        if ($Result)
        {
            Write-Host "$Process is running." -ForegroundColor Green
            $Status = "Found"
            $Found++
        }
        else
        {
            Write-Warning "$Process was not found."
            Add-Content -Path "MissingProcesses.txt" -Value $Process
            $Status = "Missing"
            $Missing++
        }

        $Results += [PSCustomObject]@{
            ProcessName = $Process
            Status      = $Status
        }
    }
    try{
    $Results | Export-Excel -Path "ProcessResults.xlsx"
    }
    catch{
        Write-Warning "Excel export failed: $($PSItem.Exception.Message)"
    }

    Write-Host "----- Summary -----" -ForegroundColor Cyan
    Write-Host "Processes checked: $($Processes.Count)" -ForegroundColor Cyan
    Write-Host "Found: $Found" -ForegroundColor Green
    Write-Host "Missing: $Missing" -ForegroundColor Red
}

Get-ProcessInfo -Processes "lsass","spoolsv","dns","w3wp"
