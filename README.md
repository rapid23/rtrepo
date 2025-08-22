rtrepo
======



param (
    [Parameter(Mandatory = $true)]
    [string]$ServerList
)

# Split comma-separated server names into an array
$servers = $ServerList -split "," | ForEach-Object { $_.Trim() }

# Output file paths
$outputAppPools = "C:\Temp\IIS_AppPools.xlsx"
$outputVirtualDirs = "C:\Temp\IIS_VirtualDirectories.xlsx"

# Arrays to store results separately
$appPoolsResults = @()
$virtualDirResults = @()

foreach ($server in $servers) {
    try {
        Write-Host "Fetching data from $server..." -ForegroundColor Cyan
        
        # Fetch Application Pools
        $appPools = Invoke-Command -ComputerName $server -ScriptBlock {
            Import-Module WebAdministration
            Get-ChildItem IIS:\AppPools | Select-Object Name, State
        }

        foreach ($pool in $appPools) {
            $appPoolsResults += [PSCustomObject]@{
                ServerName = $server
                AppPool    = $pool.Name
                State      = $pool.State
            }
        }

        # Fetch Virtual Directories
        $virtualDirs = Invoke-Command -ComputerName $server -ScriptBlock {
            Import-Module WebAdministration
            Get-WebVirtualDirectory | Select-Object Name, PhysicalPath
        }

        foreach ($vd in $virtualDirs) {
            $virtualDirResults += [PSCustomObject]@{
                ServerName     = $server
                VirtualDirName = $vd.Name
                PhysicalPath   = $vd.PhysicalPath
            }
        }
    }
    catch {
        Write-Warning "Failed to fetch data from $server. Error: $_"
    }
}

# Export results to Excel (different files)
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Force -Scope CurrentUser
}

$appPoolsResults | Export-Excel -Path $outputAppPools -AutoSize -Title "IIS App Pools"
$virtualDirResults | Export-Excel -Path $outputVirtualDirs -AutoSize -Title "IIS Virtual Directories"

Write-Host "Data collection complete." -ForegroundColor Green
Write-Host "AppPools file: $outputAppPools"
Write-Host "VirtualDirs file: $outputVirtualDirs"