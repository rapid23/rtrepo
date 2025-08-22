pipeline {
    agent { label 'windows' } // Windows Jenkins slave

    parameters {
        string(name: 'SERVER_LIST', defaultValue: '', description: 'Comma-separated list of servers')
        credentials(name: 'REMOTE_CRED', description: 'Windows credentials for remote servers', defaultValue: 'my-windows-cred')
    }

    environment {
        OUTPUT_APPPOOLS = "C:\\Temp\\IIS_AppPools.xlsx"
        OUTPUT_VDIRS   = "C:\\Temp\\IIS_VirtualDirectories.xlsx"
    }

    stages {
        stage('Collect IIS Info') {
            steps {
                script {
                    // Function to run PowerShell for IIS data collection
                    def collectIISData = { servers, credId ->
                        withCredentials([usernamePassword(credentialsId: credId, usernameVariable: 'USR', passwordVariable: 'PSW')]) {
                            def psScript = """
param([string]\$ServerList)

\$cred = New-Object System.Management.Automation.PSCredential (
    '\$env:USR', ('\$env:PSW' | ConvertTo-SecureString -AsPlainText -Force)
)

\$servers = \$ServerList -split "," | ForEach-Object { \$_ .Trim() }
\$appPoolsResults = @()
\$virtualDirResults = @()

foreach (\$server in \$servers) {
    try {
        Write-Host "Fetching data from \$server..." -ForegroundColor Cyan

        # AppPools
        \$appPools = Invoke-Command -ComputerName \$server -Credential \$cred -ScriptBlock {
            Import-Module WebAdministration
            Get-ChildItem IIS:\\AppPools | Select-Object Name, State
        }

        foreach (\$pool in \$appPools) {
            \$appPoolsResults += [PSCustomObject]@{
                ServerName = \$server
                AppPool    = \$pool.Name
                State      = \$pool.State
            }
        }

        # VirtualDirs
        \$virtualDirs = Invoke-Command -ComputerName \$server -Credential \$cred -ScriptBlock {
            Import-Module WebAdministration
            Get-WebVirtualDirectory | Select-Object Name, PhysicalPath
        }

        foreach (\$vd in \$virtualDirs) {
            \$virtualDirResults += [PSCustomObject]@{
                ServerName     = \$server
                VirtualDirName = \$vd.Name
                PhysicalPath   = \$vd.PhysicalPath
            }
        }
    }
    catch {
        Write-Warning "Failed to fetch data from \$server. Error: \$_"
    }
}

# Export results
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Force -Scope CurrentUser
}

\$appPoolsResults | Export-Excel -Path '${OUTPUT_APPPOOLS}' -AutoSize -Title "IIS App Pools"
\$virtualDirResults | Export-Excel -Path '${OUTPUT_VDIRS}' -AutoSize -Title "IIS Virtual Directories"

Write-Host "Data collection complete."
Write-Host "AppPools file: ${OUTPUT_APPPOOLS}"
Write-Host "VirtualDirs file: ${OUTPUT_VDIRS}"
"""
                            powershell(script: psScript, returnStatus: true, args: ["-ServerList", servers])
                        }
                    }

                    // Call the function
                    collectIISData(params.SERVER_LIST, 'REMOTE_CRED')
                }
            }
        }
    }

    post {
        always {
            echo "IIS data collection finished. Files saved on Jenkins slave."
        }
    }
}