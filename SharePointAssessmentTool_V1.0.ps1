<#
.SYNOPSIS
    Generates a comprehensive SharePoint health check report in HTML and CSV formats.
.DESCRIPTION
    This script collects various SharePoint information including server details, web applications,
    services, databases, and performance metrics, then exports them to an HTML report and individual CSV files.
.NOTES
    File Name      : SharePointAssessmentTool_V1.0.ps1
    Author         : Abdullah Zmaili
    Version       : 1.0
    Date Created : 2025-June-17
    Prerequisite   : PowerShell 5.1 or later, SharePoint PowerShell Module, Administrator privileges
#>

# Add SharePoint PowerShell Module
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Import-Module Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Import SQL Server module for database queries
Import-Module SqlServer -ErrorAction SilentlyContinue

# === Prompt user for the directory to save the files ===
$path = Read-Host "Enter the full path (without filename) to save the reports (e.g., C:\temp)"

# === Create directory if it doesn't exist ===
if (-not (Test-Path -Path $path)) {
    New-Item -ItemType Directory -Path $path -Force | Out-Null
}


# === Define output file paths ===
$ServerName = hostname
$htmlFile = Join-Path -Path $path -ChildPath "$ServerName-SharePointReport.html"
$csvSPServers = Join-Path -Path $path -ChildPath "$ServerName-SPServers.csv"
$csvWebApps = Join-Path -Path $path -ChildPath "$ServerName-WebApplications.csv"
$csvCentralAdmin = Join-Path -Path $path -ChildPath "$ServerName-CentralAdmin.csv"
$csvSiteCollections = Join-Path -Path $path -ChildPath "$ServerName-SiteCollections.csv"
$csvSiteAdmins = Join-Path -Path $path -ChildPath "$ServerName-SiteAdmins.csv"
$csvSiteUsers = Join-Path -Path $path -ChildPath "$ServerName-SiteUsers.csv"
$csvFarmVersion = Join-Path -Path $path -ChildPath "$ServerName-FarmVersion.csv"
$csvServerServices = Join-Path -Path $path -ChildPath "$ServerName-ServerServices.csv"
$csvDiagnosticConfig = Join-Path -Path $path -ChildPath "$ServerName-DiagnosticConfig.csv"
$csvWebBindings = Join-Path -Path $path -ChildPath "$ServerName-WebBindings.csv"
$csvTLSSettings = Join-Path -Path $path -ChildPath "$ServerName-TLSSettings.csv"
$csvSQLCounters = Join-Path -Path $path -ChildPath "$ServerName-SQLCounters.csv"
$csvSPSolutions = Join-Path -Path $path -ChildPath "$ServerName-SPSolutions.csv"
$csvSPFeatures = Join-Path -Path $path -ChildPath "$ServerName-SPFeatures.csv"

# === Additional Security & Performance CSV Files ===
$csvSPDatabases = Join-Path -Path $path -ChildPath "$ServerName-SPDatabases.csv"
$csvSPSecurity = Join-Path -Path $path -ChildPath "$ServerName-SPSecurity.csv"
$csvSPFarmAdmins = Join-Path -Path $path -ChildPath "$ServerName-SPFarmAdmins.csv"
$csvSPWebAppPolicies = Join-Path -Path $path -ChildPath "$ServerName-SPWebAppPolicies.csv"
$csvSPSearchTopology = Join-Path -Path $path -ChildPath "$ServerName-SPSearchTopology.csv"
$csvSPCacheSettings = Join-Path -Path $path -ChildPath "$ServerName-SPCacheSettings.csv"
$csvSPTimerJobs = Join-Path -Path $path -ChildPath "$ServerName-SPTimerJobs.csv"
$csvSPHealthAnalyzer = Join-Path -Path $path -ChildPath "$ServerName-SPHealthAnalyzer.csv"
$csvSPContentTypes = Join-Path -Path $path -ChildPath "$ServerName-SPContentTypes.csv"
$csvSPWebParts = Join-Path -Path $path -ChildPath "$ServerName-SPWebParts.csv"
$csvSPIISSettings = Join-Path -Path $path -ChildPath "$ServerName-SPIISSettings.csv"
$csvSPBackupHistory = Join-Path -Path $path -ChildPath "$ServerName-SPBackupHistory.csv"
$csvSPUserProfiles = Join-Path -Path $path -ChildPath "$ServerName-SPUserProfiles.csv"
$csvSPBlobCache = Join-Path -Path $path -ChildPath "$ServerName-SPBlobCache.csv"
$csvMissingUpdates = Join-Path -Path $path -ChildPath "$ServerName-MissingUpdates.csv"
$csvSPServiceAccounts = Join-Path -Path $path -ChildPath "$ServerName-SPServiceAccounts.csv"

Write-Host "`n=== SHAREPOINT HEALTH CHECK ===" -ForegroundColor Cyan
Write-Host "Starting SharePoint assessment..." -ForegroundColor Yellow
Write-Host "Output Path: $path" -ForegroundColor Green

# ----------------------------
# HELPER FUNCTIONS
# ----------------------------

function Get-TLSRegistryValue {
    <#
    .SYNOPSIS
        Helper function to safely retrieve TLS registry values
    .PARAMETER Path
        Registry path to check
    .PARAMETER Name
        Registry value name to retrieve
    #>
    param(
        [string]$Path,
        [string]$Name
    )
    
    try {
        if (Test-Path $Path) {
            $value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($value) {
                return $value.$Name
            } else {
                return "Not Set"
            }
        } else {
            return "Registry Key Not Found"
        }
    } catch {
        return "Unable to Read"
    }
}

function Get-IconHTML {
    <#
    .SYNOPSIS
        Helper function to get HTML icon symbols
    .PARAMETER IconName
        Name of the icon
    #>
    param(
        [string]$IconName
    )
    
    # Return simple HTML symbols/characters based on icon name
    switch ($IconName) {
        'chart' { return '<span class="icon-emoji">&#128202;</span>' }
        'clipboard' { return '<span class="icon-emoji">&#128203;</span>' }
        'building' { return '<span class="icon-emoji">&#127970;</span>' }
        'computer' { return '<span class="icon-emoji">&#128187;</span>' }
        'gear' { return '<span class="icon-emoji">&#9881;</span>' }
        'search' { return '<span class="icon-emoji">&#128269;</span>' }
        'globe' { return '<span class="icon-emoji">&#127760;</span>' }
        'folder' { return '<span class="icon-emoji">&#128193;</span>' }
        'users' { return '<span class="icon-emoji">&#128101;</span>' }
        'user' { return '<span class="icon-emoji">&#128100;</span>' }
        'database' { return '<span class="icon-emoji">&#128451;</span>' }
        'file' { return '<span class="icon-emoji">&#128196;</span>' }
        'lock' { return '<span class="icon-emoji">&#128274;</span>' }
        'shield' { return '<span class="icon-emoji">&#128737;</span>' }
        'crown' { return '<span class="icon-emoji">&#128081;</span>' }
        'package' { return '<span class="icon-emoji">&#128230;</span>' }
        'plug' { return '<span class="icon-emoji">&#128268;</span>' }
        'lightning' { return '<span class="icon-emoji">&#9889;</span>' }
        'wrench' { return '<span class="icon-emoji">&#128295;</span>' }
        'clock' { return '<span class="icon-emoji">&#128336;</span>' }
        'hospital' { return '<span class="icon-emoji">&#127973;</span>' }
        'refresh' { return '<span class="icon-emoji">&#128260;</span>' }
        'target' { return '<span class="icon-emoji">&#127919;</span>' }
        'folder2' { return '<span class="icon-emoji">&#128194;</span>' }
        default { return '<span class="icon-emoji">&#8226;</span>' }
    }
}

# ----------------------------
# SHAREPOINT INFORMATION FUNCTION
# ----------------------------

function Get-SharePointInformation {
    <#
    .SYNOPSIS
        Collects comprehensive SharePoint information for the health check report.
    .DESCRIPTION
        This function gathers SharePoint server details, web applications, services,
        site collections, and configuration information.
        It exports data to CSV files and returns HTML sections for the main report.
    .PARAMETER Path
        The directory path where CSV files will be saved.
    .PARAMETER ServerName
        The name of the server for file naming purposes.
    .OUTPUTS
        Returns a hashtable containing HTML sections for inclusion in the main report.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [string]$ServerName
    )
    
    Write-Host "=== COLLECTING SHAREPOINT INFORMATION ===" -ForegroundColor Cyan 
    # Initialize hashtable to store HTML sections
    $htmlSections = @{
        "SPServers" = ""
        "CentralAdmin" = ""
        "WebApplications" = ""
        "SiteCollections" = ""
        "SiteAdmins" = ""
        "SiteUsers" = ""
        "FarmVersion" = ""
        "ServerServices" = ""
        "DiagnosticConfig" = ""
        "WebBindings" = ""
        "TLSSettings" = ""
        "SQLCounters" = ""
        "SPSolutions" = ""
        "SPFeatures" = ""
        "SPDatabases" = ""
        "SPSecurity" = ""
        "SPFarmAdmins" = ""
        "SPWebAppPolicies" = ""
        "SPSearchTopology" = ""
        "SPCacheSettings" = ""
        "SPTimerJobs" = ""
        "SPHealthAnalyzer" = ""
        "SPContentTypes" = ""
        "SPWebParts" = ""
        "SPIISSettings" = ""
        "SPBackupHistory" = ""
        "SPUserProfiles" = ""
        "SPBlobCache" = ""
        "MissingUpdates" = ""
        "SPServiceAccounts" = ""
    }
    
    # === SharePoint Servers ===
    Write-Host "Collecting SharePoint Servers..." -ForegroundColor Yellow
    try {
        $spServers = Get-SPServer | Select-Object Name, Role, Status, Address
        $csvSPServers = Join-Path -Path $Path -ChildPath "$ServerName-SPServers.csv"
        $spServers | Export-Csv -Path $csvSPServers -NoTypeInformation
        $htmlSections['SPServers'] = $spServers | ConvertTo-Html -Fragment -PreContent "<h2>SharePoint Servers</h2>"
        Write-Host "SharePoint Servers - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect SharePoint Servers: $($_.Exception.Message)"
        $htmlSections['SPServers'] = "<h2>SharePoint Servers</h2><p>Error collecting SharePoint server information</p>"
    }

    # === Central Administration URL ===
    Write-Host "Collecting Central Administration URL..." -ForegroundColor Yellow
    try {
        $centralAdmin = Get-SPWebApplication -IncludeCentralAdministration | Where-Object {$_.IsAdministrationWebApplication} | Select-Object -ExpandProperty Url
        $centralAdminObj = [PSCustomObject]@{
            CentralAdminURL = $centralAdmin
            CollectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $csvCentralAdmin = Join-Path -Path $Path -ChildPath "$ServerName-CentralAdmin.csv"
        $centralAdminObj | Export-Csv -Path $csvCentralAdmin -NoTypeInformation
        $htmlSections['CentralAdmin'] = $centralAdminObj | ConvertTo-Html -Fragment -PreContent "<h2>Central Administration</h2>"
        Write-Host "Central Administration URL - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Central Administration URL: $($_.Exception.Message)"
        $htmlSections['CentralAdmin'] = "<h2>Central Administration</h2><p>Error collecting Central Administration information</p>"
    }

    # === Web Applications ===
    Write-Host "Collecting Web Applications..." -ForegroundColor Yellow
    try {
        $webApps = Get-SPWebApplication | ForEach-Object {
            $webApp = $_
            $webApp.IisSettings.Values | ForEach-Object {
                [PSCustomObject]@{
                    WebAppUrl        = $webApp.Url
                    Zone             = $_.Zone
                    ClaimsEnabled    = $webApp.UseClaimsAuthentication
                    AuthenticationMode = $_.AuthenticationMode
                    Providers        = ($_.ClaimsAuthenticationProviders | ForEach-Object { $_.DisplayName }) -join ", "
                }
            }
        }
        $csvWebApps = Join-Path -Path $Path -ChildPath "$ServerName-WebApplications.csv"
        $webApps | Export-Csv -Path $csvWebApps -NoTypeInformation
        $htmlSections['WebApplications'] = $webApps | ConvertTo-Html -Fragment -PreContent "<h2>Web Applications</h2>"
        Write-Host "Web Applications - Completed" -ForegroundColor Green
         } catch {
        Write-Warning "Failed to collect Web Applications: $($_.Exception.Message)"
        $htmlSections['WebApplications'] = "<h2>Web Applications</h2><p>Error collecting Web Applications information</p>"
         }    # === Site Collections ===
    Write-Host "Collecting Site Collections..." -ForegroundColor Yellow
    try {
        $siteCollections = Get-SPSite -Limit All | Select-Object Url, Owner, SecondaryContact, @{Name="DatabaseName"; Expression={$_.ContentDatabase.Name}}, @{Name="SizeGB"; Expression={[math]::Round($_.Usage.Storage/1GB, 2)}}
        $csvSiteCollections = Join-Path -Path $Path -ChildPath "$ServerName-SiteCollections.csv"
        $siteCollections | Export-Csv -Path $csvSiteCollections -NoTypeInformation
        $htmlSections['SiteCollections'] = $siteCollections | ConvertTo-Html -Fragment -PreContent "<h2>Site Collections</h2>"
        Write-Host "Site Collections - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Site Collections: $($_.Exception.Message)"
        $htmlSections['SiteCollections'] = "<h2>Site Collections</h2><p>Error collecting Site Collections information</p>"
    }

    # === Site Collection Administrators ===
    Write-Host "Collecting Site Collection Administrators..." -ForegroundColor Yellow
    try {
        $siteAdmins = @()
        Get-SPSite -Limit All | ForEach-Object {
            $siteUrl = $_.Url
            Write-Host "Site: $siteUrl" -ForegroundColor Gray
            
            $_.RootWeb.SiteAdministrators | ForEach-Object {
                $siteAdmins += [PSCustomObject]@{
                    SiteUrl = $siteUrl
                    UserLogin = $_.UserLogin
                    DisplayName = $_.DisplayName
                    Email = $_.Email
                    IsSiteAdmin = $_.IsSiteAdmin
                    CollectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }
        
        $csvSiteAdmins = Join-Path -Path $Path -ChildPath "$ServerName-SiteAdmins.csv"
        $siteAdmins | Export-Csv -Path $csvSiteAdmins -NoTypeInformation
        
        # Create summary for HTML (show first 20 entries)
        $siteAdminsSummary = $siteAdmins | Select-Object *
        $htmlSections['SiteAdmins'] = $siteAdminsSummary | ConvertTo-Html -Fragment -PreContent "<h2>Site Collection Administrators</h2>"
        Write-Host "Site Collection Administrators - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Site Collection Administrators: $($_.Exception.Message)"
        $htmlSections['SiteAdmins'] = "<h2>Site Collection Administrators</h2><p>Error collecting Site Collection Administrators information</p>"
    }

    # === Site Users (Summary) ===
    Write-Host "Collecting Site Users Summary..." -ForegroundColor Yellow
    try {
        $siteUsers = @()
        Get-SPWebApplication | Get-SPSite -Limit All | Get-SPWeb -Limit All | ForEach-Object {
            $web = $_
            $_.Users | ForEach-Object {
                $siteUsers += [PSCustomObject]@{
                    SiteUrl = $web.Url
                    DisplayName = $_.DisplayName
                    LoginName = $_.LoginName
                    Roles = ($_.Groups | Select-Object -ExpandProperty Name) -join ", "
                }
            }
        }
        $csvSiteUsers = Join-Path -Path $Path -ChildPath "$ServerName-SiteUsers.csv"
        $siteUsers | Export-Csv -Path $csvSiteUsers -NoTypeInformation
        
        # Create summary for HTML
        $userSummary = $siteUsers | Group-Object SiteUrl | Select-Object @{Name="SiteUrl"; Expression={$_.Name}}, @{Name="UserCount"; Expression={$_.Count}}
        $htmlSections['SiteUsers'] = $userSummary | ConvertTo-Html -Fragment -PreContent "<h2>Site Users Summary</h2>"
        Write-Host "Site Users Summary - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Site Users: $($_.Exception.Message)"
        $htmlSections['SiteUsers'] = "<h2>Site Users Summary</h2><p>Error collecting Site Users information</p>"
    }

    # === Farm Version ===
    Write-Host "Collecting Farm Version..." -ForegroundColor Yellow
    try {
        $farmVersion = (Get-SPFarm).BuildVersion
        $farmVersionObj = [PSCustomObject]@{
            FarmBuildVersion = $farmVersion.ToString()
            CollectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $csvFarmVersion = Join-Path -Path $Path -ChildPath "$ServerName-FarmVersion.csv"
        $farmVersionObj | Export-Csv -Path $csvFarmVersion -NoTypeInformation
        $htmlSections['FarmVersion'] = $farmVersionObj | ConvertTo-Html -Fragment -PreContent "<h2>Farm Version</h2>"
        Write-Host "Farm Version - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Farm Version: $($_.Exception.Message)"        $htmlSections['FarmVersion'] = "<h2>Farm Version</h2><p>Error collecting Farm Version information</p>"
    }

    # === Service Accounts ===
    Write-Host "Collecting Service Accounts..." -ForegroundColor Yellow
    try {
        $serviceAccounts = Get-SPServiceApplicationPool | Select-Object Name, ProcessAccountName, Farm
        $csvSPServiceAccounts = Join-Path -Path $Path -ChildPath "$ServerName-SPServiceAccounts.csv"
        $serviceAccounts | Export-Csv -Path $csvSPServiceAccounts -NoTypeInformation
        $htmlSections['SPServiceAccounts'] = $serviceAccounts | ConvertTo-Html -Fragment -PreContent "<h2>Service Accounts</h2>"
        Write-Host "Service Accounts - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Service Accounts: $($_.Exception.Message)"
        $htmlSections['SPServiceAccounts'] = "<h2>Service Accounts</h2><p>Error collecting Service Accounts information</p>"
    }

    # === Server Services Detail ===
    Write-Host "Collecting Server Services Detail..." -ForegroundColor Yellow
    try {
        $serverServices = @()
        Get-SPServer | ForEach-Object {
            $server = $_.Name
            Get-SPServiceInstance -Server $_ | ForEach-Object {
                $serverServices += [PSCustomObject]@{
                    ServerName = $server
                    ServiceType = $_.TypeName
                    Status = $_.Status
                    DisplayName = $_.DisplayName
                }
            }
        }
        $csvServerServices = Join-Path -Path $Path -ChildPath "$ServerName-ServerServices.csv"
        $serverServices | Export-Csv -Path $csvServerServices -NoTypeInformation
        $htmlSections['ServerServices'] = $serverServices | ConvertTo-Html -Fragment -PreContent "<h2>Server Services Detail</h2>"
        Write-Host "Server Services Detail - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Server Services Detail: $($_.Exception.Message)"
        $htmlSections['ServerServices'] = "<h2>Server Services Detail</h2><p>Error collecting Server Services information</p>"
    }

    # === Diagnostic Configuration ===
    Write-Host "Collecting Diagnostic Configuration..." -ForegroundColor Yellow
    try {
        $diagnosticConfig = Get-SPDiagnosticConfig | Select-Object LogLocation, EventLogFloodProtectionEnabled, DaysToKeepLogs, LogCutInterval
        $csvDiagnosticConfig = Join-Path -Path $Path -ChildPath "$ServerName-DiagnosticConfig.csv"
        $diagnosticConfig | Export-Csv -Path $csvDiagnosticConfig -NoTypeInformation
        $htmlSections['DiagnosticConfig'] = $diagnosticConfig | ConvertTo-Html -Fragment -PreContent "<h2>Diagnostic Configuration</h2>"
        Write-Host "Diagnostic Configuration - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Diagnostic Configuration: $($_.Exception.Message)"
        $htmlSections['DiagnosticConfig'] = "<h2>Diagnostic Configuration</h2><p>Error collecting Diagnostic Configuration information</p>"
    }

    # === Web Bindings (HTTPS) ===
    Write-Host "Collecting Web Bindings..." -ForegroundColor Yellow
    try {
        Import-Module WebAdministration -ErrorAction SilentlyContinue
        $webBindings = Get-WebBinding | Where-Object { $_.protocol -eq "https" } | Select-Object BindingInformation, protocol, @{Name="SiteName"; Expression={(Get-Website | Where-Object {$_.Id -eq $_.ItemXPath.Split("'")[1]}).Name}}
        $csvWebBindings = Join-Path -Path $Path -ChildPath "$ServerName-WebBindings.csv"
        $webBindings | Export-Csv -Path $csvWebBindings -NoTypeInformation
        $htmlSections['WebBindings'] = $webBindings | ConvertTo-Html -Fragment -PreContent "<h2>HTTPS Web Bindings</h2>"
        Write-Host "Web Bindings - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Web Bindings: $($_.Exception.Message)"
        $htmlSections['WebBindings'] = "<h2>HTTPS Web Bindings</h2><p>Error collecting Web Bindings information</p>"
    }    # === TLS Settings ===
    Write-Host "Collecting TLS Settings..." -ForegroundColor Yellow
    try {
        $regSettings = @()
          # Check .NET Framework TLS settings (32-bit)
        $regKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319'
        $regSettings += [PSCustomObject]@{
            Category = ".NET Framework 32-bit"
            Path = $regKey
            Name = 'SystemDefaultTlsVersions'
            Value = Get-TLSRegistryValue -Path $regKey -Name 'SystemDefaultTlsVersions'
            Description = "Use system default TLS versions for .NET Framework 32-bit applications"
        }
        $regSettings += [PSCustomObject]@{
            Category = ".NET Framework 32-bit"
            Path = $regKey
            Name = 'SchUseStrongCrypto'
            Value = Get-TLSRegistryValue -Path $regKey -Name 'SchUseStrongCrypto'
            Description = "Use strong cryptography for .NET Framework 32-bit applications"
        }

        # Check .NET Framework TLS settings (64-bit)
        $regKey = 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319'
        $regSettings += [PSCustomObject]@{
            Category = ".NET Framework 64-bit"
            Path = $regKey
            Name = 'SystemDefaultTlsVersions'
            Value = Get-TLSRegistryValue -Path $regKey -Name 'SystemDefaultTlsVersions'
            Description = "Use system default TLS versions for .NET Framework 64-bit applications"
        }
        $regSettings += [PSCustomObject]@{
            Category = ".NET Framework 64-bit"
            Path = $regKey
            Name = 'SchUseStrongCrypto'
            Value = Get-TLSRegistryValue -Path $regKey -Name 'SchUseStrongCrypto'
            Description = "Use strong cryptography for .NET Framework 64-bit applications"
        }        # Check TLS 1.2 Server settings
        $regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server'
        $regSettings += [PSCustomObject]@{
            Category = "TLS 1.2 Server"
            Path = $regKey
            Name = 'Enabled'
            Value = Get-TLSRegistryValue -Path $regKey -Name 'Enabled'
            Description = "Enable TLS 1.2 for server-side connections"
        }
        $regSettings += [PSCustomObject]@{
            Category = "TLS 1.2 Server"
            Path = $regKey
            Name = 'DisabledByDefault'
            Value = Get-TLSRegistryValue -Path $regKey -Name 'DisabledByDefault'
            Description = "Disable TLS 1.2 by default for server-side connections"
        }

        # Check TLS 1.2 Client settings
        $regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client'
        $regSettings += [PSCustomObject]@{
            Category = "TLS 1.2 Client"
            Path = $regKey
            Name = 'Enabled'
            Value = Get-TLSRegistryValue -Path $regKey -Name 'Enabled'
            Description = "Enable TLS 1.2 for client-side connections"
        }
        $regSettings += [PSCustomObject]@{
            Category = "TLS 1.2 Client"
            Path = $regKey
            Name = 'DisabledByDefault'
            Value = Get-TLSRegistryValue -Path $regKey -Name 'DisabledByDefault'
            Description = "Disable TLS 1.2 by default for client-side connections"
        }

        # Check for older TLS versions that should be disabled
        $oldTlsVersions = @("SSL 2.0", "SSL 3.0", "TLS 1.0", "TLS 1.1")
        foreach ($version in $oldTlsVersions) {
            $serverKey = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$version\Server"
            $clientKey = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$version\Client"
              $regSettings += [PSCustomObject]@{
                Category = "$version Server (Legacy)"
                Path = $serverKey
                Name = 'Enabled'
                Value = Get-TLSRegistryValue -Path $serverKey -Name 'Enabled'
                Description = "Enable $version for server-side connections (Legacy Protocol)"
            }
            
            $regSettings += [PSCustomObject]@{
                Category = "$version Client (Legacy)"
                Path = $clientKey
                Name = 'Enabled'
                Value = Get-TLSRegistryValue -Path $clientKey -Name 'Enabled'
                Description = "Enable $version for client-side connections (Legacy Protocol)"
            }
        }

        # Check TLS 1.3 settings (if supported)
        $tls13ServerKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server'
        $tls13ClientKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Client'
          $regSettings += [PSCustomObject]@{
            Category = "TLS 1.3 Server"
            Path = $tls13ServerKey
            Name = 'Enabled'
            Value = Get-TLSRegistryValue -Path $tls13ServerKey -Name 'Enabled'
            Description = "Enable TLS 1.3 for server-side connections"
        }
        
        $regSettings += [PSCustomObject]@{
            Category = "TLS 1.3 Client"
            Path = $tls13ClientKey
            Name = 'Enabled'
            Value = Get-TLSRegistryValue -Path $tls13ClientKey -Name 'Enabled'
            Description = "Enable TLS 1.3 for client-side connections"
        }

        $csvTLSSettings = Join-Path -Path $Path -ChildPath "$ServerName-TLSSettings.csv"
        $regSettings | Export-Csv -Path $csvTLSSettings -NoTypeInformation
        
        # Create summary for HTML (show key settings)
        $tlsSummary = $regSettings | Where-Object {$_.Category -like "*TLS 1.2*" -or $_.Category -like "*.NET Framework*"}
        $htmlSections['TLSSettings'] = $tlsSummary | ConvertTo-Html -Fragment -PreContent "<h2>TLS Security Configuration</h2>"
        Write-Host "TLS Settings - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect TLS Settings: $($_.Exception.Message)"
        $htmlSections['TLSSettings'] = "<h2>TLS Security Configuration</h2><p>Error collecting TLS Settings information</p>"
    }

    # === SQL Performance Counters ===
    Write-Host "Collecting SQL Performance Counters..." -ForegroundColor Yellow
    try {
        $sqlCounters = @()
        
        # Page Life Expectancy
        try {
            $pageLifeExpectancy = Get-Counter '\SQLServer:Buffer Manager\Page life expectancy' -ErrorAction SilentlyContinue
            $sqlCounters += [PSCustomObject]@{
                CounterName = "Page Life Expectancy"
                Value = $pageLifeExpectancy.CounterSamples.CookedValue
                Unit = "Seconds"
                Timestamp = $pageLifeExpectancy.Timestamp
            }
        } catch {
            $sqlCounters += [PSCustomObject]@{
                CounterName = "Page Life Expectancy"
                Value = "Not Available"
                Unit = "Seconds"
                Timestamp = Get-Date
            }
        }

        # User Connections
        try {
            $userConnections = Get-Counter '\SQLServer:General Statistics\User Connections' -ErrorAction SilentlyContinue
            $sqlCounters += [PSCustomObject]@{
                CounterName = "User Connections"
                Value = $userConnections.CounterSamples.CookedValue
                Unit = "Connections"
                Timestamp = $userConnections.Timestamp
            }
        } catch {
            $sqlCounters += [PSCustomObject]@{
                CounterName = "User Connections"
                Value = "Not Available"
                Unit = "Connections"
                Timestamp = Get-Date
            }
        }

        # Disk Read Performance
        try {
            $diskRead = Get-Counter '\PhysicalDisk(_Total)\Avg. Disk sec/Read' -ErrorAction SilentlyContinue
            $sqlCounters += [PSCustomObject]@{
                CounterName = "Avg. Disk sec/Read"
                Value = [math]::Round($diskRead.CounterSamples.CookedValue, 4)
                Unit = "Seconds"
                Timestamp = $diskRead.Timestamp
            }
        } catch {
            $sqlCounters += [PSCustomObject]@{
                CounterName = "Avg. Disk sec/Read"
                Value = "Not Available"
                Unit = "Seconds"
                Timestamp = Get-Date
            }
        }

        $csvSQLCounters = Join-Path -Path $Path -ChildPath "$ServerName-SQLCounters.csv"
        $sqlCounters | Export-Csv -Path $csvSQLCounters -NoTypeInformation
        $htmlSections['SQLCounters'] = $sqlCounters | ConvertTo-Html -Fragment -PreContent "<h2>SQL Performance Counters</h2>"
        Write-Host "SQL Performance Counters - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect SQL Performance Counters: $($_.Exception.Message)"
        $htmlSections['SQLCounters'] = "<h2>SQL Performance Counters</h2><p>Error collecting SQL Performance Counters information</p>"
    }

    # === SharePoint Solutions ===
    Write-Host "Collecting SharePoint Solutions..." -ForegroundColor Yellow
    try {
        $spSolutions = Get-SPSolution | Select-Object DisplayName, Deployed, SolutionId, @{Name="DeployedServers"; Expression={($_.DeployedServers | Select-Object -ExpandProperty Name) -join ", "}}, @{Name="ContainsGlobalAssembly"; Expression={$_.ContainsGlobalAssembly}}
        $csvSPSolutions = Join-Path -Path $Path -ChildPath "$ServerName-SPSolutions.csv"
        $spSolutions | Export-Csv -Path $csvSPSolutions -NoTypeInformation
        $htmlSections['SPSolutions'] = $spSolutions | ConvertTo-Html -Fragment -PreContent "<h2>SharePoint Solutions</h2>"
        Write-Host "SharePoint Solutions - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect SharePoint Solutions: $($_.Exception.Message)"
        $htmlSections['SPSolutions'] = "<h2>SharePoint Solutions</h2><p>Error collecting SharePoint Solutions information</p>"
    }    # === SharePoint Features ===
    Write-Host "Collecting SharePoint Features..." -ForegroundColor Yellow
    try {
        $spFeatures = Get-SPFeature | Sort-Object DisplayName | Select-Object DisplayName, Id, Scope, @{Name="Activated"; Expression={$_.Status -eq "Online"}}
        $csvSPFeatures = Join-Path -Path $Path -ChildPath "$ServerName-SPFeatures.csv"
        $spFeatures | Export-Csv -Path $csvSPFeatures -NoTypeInformation
        
        # Create summary for HTML 
        $featuresSummary = $spFeatures | Select-Object *
        $htmlSections['SPFeatures'] = $featuresSummary | ConvertTo-Html -Fragment -PreContent "<h2>SharePoint Features</h2>"
        Write-Host "SharePoint Features - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect SharePoint Features: $($_.Exception.Message)"
        $htmlSections['SPFeatures'] = "<h2>SharePoint Features</h2><p>Error collecting SharePoint Features information</p>"
    }

    # === SharePoint Content Databases ===
    Write-Host "Collecting SharePoint Content Databases..." -ForegroundColor Yellow
    try {
        $spDatabases = Get-SPContentDatabase | Select-Object Name, Server, @{Name="SizeGB"; Expression={[math]::Round($_.DiskSizeRequired/1GB, 2)}}, 
            @{Name="SiteCount"; Expression={$_.CurrentSiteCount}}, 
            @{Name="MaxSiteCount"; Expression={$_.MaximumSiteCount}}, 
            Status, 
            @{Name="LastBackup"; Expression={
                try {
                    $backupInfo = Invoke-Sqlcmd -ServerInstance $_.Server -Database $_.Name -Query "SELECT TOP 1 backup_start_date FROM msdb.dbo.backupset WHERE database_name = '$($_.Name)' ORDER BY backup_start_date DESC" -ErrorAction SilentlyContinue
                    if ($backupInfo) { $backupInfo.backup_start_date } else { "No backup found" }
                } catch { "Unable to check" }
            }}
        $csvSPDatabases = Join-Path -Path $Path -ChildPath "$ServerName-SPDatabases.csv"
        $spDatabases | Export-Csv -Path $csvSPDatabases -NoTypeInformation
        $htmlSections['SPDatabases'] = $spDatabases | ConvertTo-Html -Fragment -PreContent "<h2>SharePoint Content Databases</h2>"
        Write-Host "SharePoint Content Databases - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect SharePoint Content Databases: $($_.Exception.Message)"
        $htmlSections['SPDatabases'] = "<h2>SharePoint Content Databases</h2><p>Error collecting Content Databases information</p>"
    }

    # === SharePoint Security Configuration ===
    Write-Host "Collecting SharePoint Security Configuration..." -ForegroundColor Yellow
    try {
        $spSecurity = @()
        
        # Authentication providers
        Get-SPWebApplication | ForEach-Object {
            $webApp = $_
            $webApp.IisSettings.Values | ForEach-Object {
                $spSecurity += [PSCustomObject]@{
                    Category = "Authentication"
                    WebApplication = $webApp.Url
                    Zone = $_.Zone
                    AuthenticationMode = if ($webApp.UseClaimsAuthentication) { "Claims" } else { "Classic" }
                    AnonymousAccess = $_.AllowAnonymous
                    WindowsAuth = $_.UseWindowsClaimsAuthenticationProvider
                    FormsAuth = $_.UseForms
                    TrustedProvider = $_.TrustedIdentityProviders.Count
                }
            }
        }
        
        # Blocked file types
        Get-SPWebApplication | ForEach-Object {
            $blockedFiles = $_.BlockedFileExtensions -join ", "
            $spSecurity += [PSCustomObject]@{
                Category = "File Security"
                WebApplication = $_.Url
                Zone = "All"
                AuthenticationMode = "N/A"
                AnonymousAccess = "N/A"
                WindowsAuth = "N/A"
                FormsAuth = "N/A"
                TrustedProvider = "N/A"
                Details = "Blocked Extensions: $blockedFiles"
            }
        }
        
        $csvSPSecurity = Join-Path -Path $Path -ChildPath "$ServerName-SPSecurity.csv"
        $spSecurity | Export-Csv -Path $csvSPSecurity -NoTypeInformation
        $htmlSections['SPSecurity'] = $spSecurity | ConvertTo-Html -Fragment -PreContent "<h2>SharePoint Security Configuration</h2>"
        Write-Host "SharePoint Security Configuration - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect SharePoint Security Configuration: $($_.Exception.Message)"
        $htmlSections['SPSecurity'] = "<h2>SharePoint Security Configuration</h2><p>Error collecting Security Configuration information</p>"
    }    # === Farm Administrators ===
    Write-Host "Collecting Farm Administrators..." -ForegroundColor Yellow
    try {
        $farmAdminsList = @()
        
        # Method 1: Try using SharePoint Central Administration Security
        try {
            $caWebApp = Get-SPWebApplication -IncludeCentralAdministration | Where-Object {$_.IsAdministrationWebApplication}
            if ($caWebApp) {
                $caWeb = $caWebApp.Sites[0].RootWeb
                $farmAdminsGroup = $caWeb.SiteGroups | Where-Object {$_.Name -like "*Farm Administrators*"}
                if ($farmAdminsGroup) {
                    foreach ($user in $farmAdminsGroup.Users) {
                        $farmAdminsList += [PSCustomObject]@{
                            PrincipalName = $user.LoginName
                            DisplayName = $user.Name
                            Type = if ($user.LoginName -like "*\*") { "Windows Account" } else { "Claims Account" }
                            Source = "Central Admin Group"
                            CollectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        }
                    }
                }
            }
        } catch {
            Write-Warning "Failed to get farm admins from Central Admin: $($_.Exception.Message)"
        }
        
        # Method 2: Try using Configuration Database query (if SQL access is available)
        if ($farmAdminsList.Count -eq 0) {
            try {
                $configDB = Get-SPDatabase | Where-Object {$_.Name -like "*SharePoint_Config*" -or $_.TypeName -eq "Configuration Database"}
                if ($configDB) {
                    $farmAdminUsers = Invoke-Sqlcmd -ServerInstance $configDB.Server -Database $configDB.Name -Query "SELECT principalname FROM dbo.FarmAdministrators" -ErrorAction SilentlyContinue
                    foreach ($admin in $farmAdminUsers) {
                        $farmAdminsList += [PSCustomObject]@{
                            PrincipalName = $admin.principalname
                            DisplayName = $admin.principalname
                            Type = if ($admin.principalname -like "*\*") { "Windows Account" } else { "Claims Account" }
                            Source = "Configuration Database"
                            CollectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        }
                    }
                }
            } catch {
                Write-Warning "Failed to query configuration database: $($_.Exception.Message)"
            }
        }
        
        # Method 3: Alternative approach using SPSecurity
        if ($farmAdminsList.Count -eq 0) {
            try {
                $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                $isCurrentUserFarmAdmin = Get-SPFarm | Select-Object -ExpandProperty CurrentUserIsAdmin
                
                $farmAdminsList += [PSCustomObject]@{
                    PrincipalName = $currentUser
                    DisplayName = $currentUser
                    Type = "Current User"
                    Source = "Current Context (Farm Admin: $isCurrentUserFarmAdmin)"
                    CollectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
                
                # Try to get farm administrators using PowerShell security context
                $farm = Get-SPFarm
                if ($farm.Properties.ContainsKey("FarmAdministrators")) {
                    $adminAccounts = $farm.Properties["FarmAdministrators"] -split ";"
                    foreach ($admin in $adminAccounts) {
                        if ($admin -and $admin.Trim() -ne "") {
                            $farmAdminsList += [PSCustomObject]@{
                                PrincipalName = $admin.Trim()
                                DisplayName = $admin.Trim()
                                Type = if ($admin -like "*\*") { "Windows Account" } else { "Claims Account" }
                                Source = "Farm Properties"
                                CollectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            }
                        }
                    }
                }
            } catch {
                Write-Warning "Failed to get farm admin using alternative method: $($_.Exception.Message)"
            }
        }
        
        # If still no results, provide a helpful message
        if ($farmAdminsList.Count -eq 0) {
            $farmAdminsList += [PSCustomObject]@{
                PrincipalName = "Unable to retrieve farm administrators"
                DisplayName = "Check manually via Central Administration > Security > Manage the farm administrators group"
                Type = "Information"
                Source = "Manual Check Required"
                CollectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        $csvSPFarmAdmins = Join-Path -Path $Path -ChildPath "$ServerName-SPFarmAdmins.csv"
        $farmAdminsList | Export-Csv -Path $csvSPFarmAdmins -NoTypeInformation
        $htmlSections['SPFarmAdmins'] = $farmAdminsList | ConvertTo-Html -Fragment -PreContent "<h2>Farm Administrators</h2>"
        Write-Host "Farm Administrators - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Farm Administrators: $($_.Exception.Message)"
        $htmlSections['SPFarmAdmins'] = "<h2>Farm Administrators</h2><p>Error collecting Farm Administrators information</p>"
    }

    # === Web Application Policies ===
    Write-Host "Collecting Web Application Policies..." -ForegroundColor Yellow
    try {
        $webAppPolicies = @()
        Get-SPWebApplication | ForEach-Object {
            $webApp = $_
            $_.Policies | ForEach-Object {
                $webAppPolicies += [PSCustomObject]@{
                    WebApplication = $webApp.Url
                    UserName = $_.UserName
                    DisplayName = $_.DisplayName
                    PolicyRoles = ($_.PolicyRoleBindings | Select-Object -ExpandProperty Name) -join ", "
                    IsSystemUser = $_.IsSystemUser
                }
            }
        }
        $csvSPWebAppPolicies = Join-Path -Path $Path -ChildPath "$ServerName-SPWebAppPolicies.csv"
        $webAppPolicies | Export-Csv -Path $csvSPWebAppPolicies -NoTypeInformation
        $htmlSections['SPWebAppPolicies'] = $webAppPolicies | ConvertTo-Html -Fragment -PreContent "<h2>Web Application Policies</h2>"
        Write-Host "Web Application Policies - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Web Application Policies: $($_.Exception.Message)"
        $htmlSections['SPWebAppPolicies'] = "<h2>Web Application Policies</h2><p>Error collecting Web Application Policies information</p>"
    }

    # === Search Service Topology ===
    Write-Host "Collecting Search Service Topology..." -ForegroundColor Yellow
    try {
        $searchTopology = @()
        $searchServiceApp = Get-SPEnterpriseSearchServiceApplication -ErrorAction SilentlyContinue
        if ($searchServiceApp) {
            Get-SPEnterpriseSearchTopology -SearchApplication $searchServiceApp | ForEach-Object {
                $topology = $_
                Get-SPEnterpriseSearchComponent -SearchTopology $topology | ForEach-Object {
                    $searchTopology += [PSCustomObject]@{
                        TopologyId = $topology.TopologyId
                        ComponentId = $_.ComponentId
                        Name = $_.Name
                        ServerName = $_.ServerName
                        IndexPartitionOrdinal = $_.IndexPartitionOrdinal
                        State = $_.State
                    }
                }
            }
        } else {
            $searchTopology += [PSCustomObject]@{
                TopologyId = "N/A"
                ComponentId = "N/A"
                Name = "No Search Service Application Found"
                ServerName = "N/A"
                IndexPartitionOrdinal = "N/A"
                State = "N/A"
            }
        }
        $csvSPSearchTopology = Join-Path -Path $Path -ChildPath "$ServerName-SPSearchTopology.csv"
        $searchTopology | Export-Csv -Path $csvSPSearchTopology -NoTypeInformation
        $htmlSections['SPSearchTopology'] = $searchTopology | ConvertTo-Html -Fragment -PreContent "<h2>Search Service Topology</h2>"
        Write-Host "Search Service Topology - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Search Service Topology: $($_.Exception.Message)"
        $htmlSections['SPSearchTopology'] = "<h2>Search Service Topology</h2><p>Error collecting Search Service Topology information</p>"
    }

    # === Cache Settings ===
    Write-Host "Collecting Cache Settings..." -ForegroundColor Yellow
    try {
        $cacheSettings = @()
        Get-SPWebApplication | ForEach-Object {
            $webApp = $_
            $cacheSettings += [PSCustomObject]@{
                WebApplication = $webApp.Url
                ObjectCacheEnabled = $webApp.Properties["object-cache-enabled"]
                ObjectCacheMaxSize = $webApp.Properties["object-cache-max-size"]
                OutputCacheEnabled = $webApp.Properties["output-cache-enabled"]
                PageOutputCacheEnabled = $webApp.Properties["page-output-cache-enabled"]
            }
        }
        $csvSPCacheSettings = Join-Path -Path $Path -ChildPath "$ServerName-SPCacheSettings.csv"
        $cacheSettings | Export-Csv -Path $csvSPCacheSettings -NoTypeInformation
        $htmlSections['SPCacheSettings'] = $cacheSettings | ConvertTo-Html -Fragment -PreContent "<h2>Cache Settings</h2>"
        Write-Host "Cache Settings - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Cache Settings: $($_.Exception.Message)"
        $htmlSections['SPCacheSettings'] = "<h2>Cache Settings</h2><p>Error collecting Cache Settings information</p>"
    }

    # === Enabled Timer Jobs ===
    Write-Host "Collecting Enabled Timer Jobs..." -ForegroundColor Yellow
    try {
        $timerJobs = Get-SPTimerJob | Where-Object {$_.IsDisabled -eq $false} | Select-Object Name, 
            @{Name="WebApplication"; Expression={if ($_.WebApplication) {$_.WebApplication.Url} else {"Farm Level"}}}, 
            @{Name="LastRunTime"; Expression={$_.LastRunTime}},
            @{Name="Schedule"; Expression={$_.Schedule.ToString()}}
        $csvSPTimerJobs = Join-Path -Path $Path -ChildPath "$ServerName-SPTimerJobs.csv"
        $timerJobs | Export-Csv -Path $csvSPTimerJobs -NoTypeInformation

        # Show all for HTML
        $timerJobsSummary = $timerJobs | Select-Object *
        $htmlSections['SPTimerJobs'] = $timerJobsSummary | ConvertTo-Html -Fragment -PreContent "<h2>Enabled Timer Jobs</h2>"
        Write-Host "Enabled Timer Jobs - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect enabled Timer Jobs: $($_.Exception.Message)"
        $htmlSections['SPTimerJobs'] = "<h2>Enabled Timer Jobs</h2><p>Error collecting enabled Timer Jobs information</p>"
    }

    # === Health Analyzer Reports ===
    Write-Host "Collecting Health Analyzer Reports..." -ForegroundColor Yellow
    try {
        $healthRules = Get-SPHealthAnalysisRule | Select-Object DisplayName, 
            Category, 
            Enabled, 
            @{Name="Severity"; Expression={$_.Severity.ToString()}},
            @{Name="Schedule"; Expression={$_.Schedule.ToString()}},
            @{Name="LastRunTime"; Expression={
                $ruleId = $_.Id
                $reports = Get-SPHealthReport | Where-Object {$_.HealthRuleId -eq $ruleId}
                if ($reports) { ($reports | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated }
                else { "Never Run" }
            }}
        $csvSPHealthAnalyzer = Join-Path -Path $Path -ChildPath "$ServerName-SPHealthAnalyzer.csv"
        $healthRules | Export-Csv -Path $csvSPHealthAnalyzer -NoTypeInformation
        $htmlSections['SPHealthAnalyzer'] = $healthRules | ConvertTo-Html -Fragment -PreContent "<h2>Health Analyzer Rules</h2>"
        Write-Host "Health Analyzer Reports - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Health Analyzer Reports: $($_.Exception.Message)"
        $htmlSections['SPHealthAnalyzer'] = "<h2>Health Analyzer Rules</h2><p>Error collecting Health Analyzer information</p>"
    }

    # === Content Types ===
    Write-Host "Collecting Content Types Summary..." -ForegroundColor Yellow
    try {
        $contentTypes = @()
        Get-SPSite -Limit All | ForEach-Object {
            $site = $_
            $_.RootWeb.ContentTypes | ForEach-Object {
                $contentTypes += [PSCustomObject]@{
                    SiteUrl = $site.Url
                    ContentTypeName = $_.Name
                    ContentTypeId = $_.Id.ToString()
                    Group = $_.Group
                    Hidden = $_.Hidden
                    Sealed = $_.Sealed
                }
                if ($contentTypes.Count -ge 100) { return } # Limit to first 100 for performance
            }
        }
        $csvSPContentTypes = Join-Path -Path $Path -ChildPath "$ServerName-SPContentTypes.csv"
        $contentTypes | Export-Csv -Path $csvSPContentTypes -NoTypeInformation
        
        # Summary for HTML
        $contentTypesSummary = $contentTypes | Group-Object Group | Select-Object @{Name="Group"; Expression={$_.Name}}, @{Name="Count"; Expression={$_.Count}}
        $htmlSections['SPContentTypes'] = $contentTypesSummary | ConvertTo-Html -Fragment -PreContent "<h2>Content Types Summary</h2>"
        Write-Host "Content Types Summary - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect Content Types: $($_.Exception.Message)"
        $htmlSections['SPContentTypes'] = "<h2>Content Types Summary</h2><p>Error collecting Content Types information</p>"
    }

    # === IIS Settings ===
    Write-Host "Collecting IIS Settings..." -ForegroundColor Yellow
    try {
        Import-Module WebAdministration -ErrorAction SilentlyContinue
        $iisSettings = @()
        Get-Website | Where-Object {$_.Name -like "*SharePoint*" -or $_.Name -like "*Central Administration*"} | ForEach-Object {
            $site = $_
            $iisSettings += [PSCustomObject]@{
                SiteName = $site.Name
                State = $site.State
                PhysicalPath = $site.PhysicalPath
                ApplicationPool = $site.ApplicationPool
                Bindings = ($site.Bindings.Collection | ForEach-Object {"$($_.protocol)://$($_.bindingInformation)"}) -join ", "
                LogFormat = $site.LogFile.LogFormat
                LogDirectory = $site.LogFile.Directory
            }
        }
        $csvSPIISSettings = Join-Path -Path $Path -ChildPath "$ServerName-SPIISSettings.csv"
        $iisSettings | Export-Csv -Path $csvSPIISSettings -NoTypeInformation
        $htmlSections['SPIISSettings'] = $iisSettings | ConvertTo-Html -Fragment -PreContent "<h2>IIS Settings</h2>"
        Write-Host "IIS Settings - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect IIS Settings: $($_.Exception.Message)"
        $htmlSections['SPIISSettings'] = "<h2>IIS Settings</h2><p>Error collecting IIS Settings information</p>"    }
    
    # === Missing Windows Updates ===
    Write-Host "Checking for Missing Windows Updates..." -ForegroundColor Yellow
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        $missingUpdatesResult = $updateSearcher.Search("IsInstalled=0")
        
        if ($missingUpdatesResult.Updates.Count -gt 0) {
            $missingUpdates = $missingUpdatesResult.Updates | ForEach-Object {
                [PSCustomObject]@{
                    Title = $_.Title
                    KB = if ($_.KBArticleIDs.Count -gt 0) { ($_.KBArticleIDs -join ", ") } else { "N/A" }
                    SizeMB = [math]::Round($_.MaxDownloadSize / 1MB, 2)
                    Severity = $_.MsrcSeverity
                    RebootRequired = $_.RebootRequired
                    ReleaseDate = if ($_.LastDeploymentChangeTime) { $_.LastDeploymentChangeTime.ToString("yyyy-MM-dd") } else { "N/A" }
                }
            }
        } else {
            $missingUpdates = @([PSCustomObject]@{
                Title = "No missing updates found"
                KB = "N/A"
                SizeMB = 0
                Severity = "N/A"
                RebootRequired = $false
                ReleaseDate = "N/A"
            })
        }
        
        $csvMissingUpdates = Join-Path -Path $Path -ChildPath "$ServerName-MissingUpdates.csv"
        $missingUpdates | Export-Csv -Path $csvMissingUpdates -NoTypeInformation
        $htmlSections['MissingUpdates'] = $missingUpdates | ConvertTo-Html -Fragment -PreContent "<h2>Missing Windows Updates</h2>"
        Write-Host "Missing Windows Updates - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to check for Missing Windows Updates: $($_.Exception.Message)"
        $htmlSections['MissingUpdates'] = "<h2>Missing Windows Updates</h2><p>Error checking for missing updates: $($_.Exception.Message)</p>"
    }

    # === User Profile Service ===
    Write-Host "Collecting User Profile Service Information..." -ForegroundColor Yellow
    try {
        $userProfileInfo = @()
        $userProfileService = Get-SPServiceApplication | Where-Object {$_.TypeName -like "*User Profile*"}
        if ($userProfileService) {
            $userProfileInfo += [PSCustomObject]@{
                ServiceName = $userProfileService.DisplayName
                Status = $userProfileService.Status
                ApplicationPool = $userProfileService.ApplicationPool.Name
                DatabaseName = if ($userProfileService.Databases) { ($userProfileService.Databases | Select-Object -ExpandProperty Name) -join ", " } else { "N/A" }
                ProfileCount = "N/A" # This would require more complex querying
                SyncConnectionCount = "N/A" # This would require more complex querying
            }
        } else {
            $userProfileInfo += [PSCustomObject]@{
                ServiceName = "User Profile Service Application Not Found"
                Status = "N/A"
                ApplicationPool = "N/A"
                DatabaseName = "N/A"
                ProfileCount = "N/A"
                SyncConnectionCount = "N/A"
            }
        }
        $csvSPUserProfiles = Join-Path -Path $Path -ChildPath "$ServerName-SPUserProfiles.csv"
        $userProfileInfo | Export-Csv -Path $csvSPUserProfiles -NoTypeInformation
        $htmlSections['SPUserProfiles'] = $userProfileInfo | ConvertTo-Html -Fragment -PreContent "<h2>User Profile Service</h2>"
        Write-Host "User Profile Service Information - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect User Profile Service Information: $($_.Exception.Message)"
        $htmlSections['SPUserProfiles'] = "<h2>User Profile Service</h2><p>Error collecting User Profile Service information</p>"
    }

    # === BLOB Cache Settings ===
    Write-Host "Collecting BLOB Cache Settings..." -ForegroundColor Yellow
    try {
        $blobCacheSettings = @()
        Get-SPWebApplication | ForEach-Object {
            $webApp = $_
            $webApp.IisSettings.Values | ForEach-Object {
                $webConfig = "$($_.Path)\web.config"
                if (Test-Path $webConfig) {
                    try {
                        [xml]$webConfigXml = Get-Content $webConfig
                        $blobCacheNode = $webConfigXml.configuration.SharePoint.BlobCache
                        $blobCacheSettings += [PSCustomObject]@{
                            WebApplication = $webApp.Url
                            Zone = $_.Zone
                            BlobCacheEnabled = if ($blobCacheNode) { $blobCacheNode.enabled } else { "Not Configured" }
                            Location = if ($blobCacheNode) { $blobCacheNode.location } else { "N/A" }
                            MaxSize = if ($blobCacheNode) { $blobCacheNode.maxSize } else { "N/A" }
                            FileTypes = if ($blobCacheNode) { $blobCacheNode.path } else { "N/A" }
                        }
                    } catch {
                        $blobCacheSettings += [PSCustomObject]@{
                            WebApplication = $webApp.Url
                            Zone = $_.Zone
                            BlobCacheEnabled = "Error Reading Config"
                            Location = "N/A"
                            MaxSize = "N/A"
                            FileTypes = "N/A"
                        }
                    }
                }
            }
        }
        $csvSPBlobCache = Join-Path -Path $Path -ChildPath "$ServerName-SPBlobCache.csv"
        $blobCacheSettings | Export-Csv -Path $csvSPBlobCache -NoTypeInformation
        $htmlSections['SPBlobCache'] = $blobCacheSettings | ConvertTo-Html -Fragment -PreContent "<h2>BLOB Cache Settings</h2>"
        Write-Host "BLOB Cache Settings - Completed" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to collect BLOB Cache Settings: $($_.Exception.Message)"
        $htmlSections['SPBlobCache'] = "<h2>BLOB Cache Settings</h2><p>Error collecting BLOB Cache Settings information</p>"
    }
    
    # === Combine Patch Management HTML Sections ===
    Write-Host "Combining Patch Management sections..." -ForegroundColor Yellow
    try {
        # Keep individual sections separate for collapsible display
        # Farm Version and Missing Updates will be displayed individually
        Write-Host "Patch Management sections prepared successfully" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to prepare Patch Management sections: $($_.Exception.Message)"
    }

    Write-Host "=== SHAREPOINT INFORMATION COLLECTION COMPLETED ===" -ForegroundColor Green
    return $htmlSections
}

# ----------------------------
# MAIN EXECUTION
# ----------------------------

Write-Host "Starting SharePoint Information Collection..." -ForegroundColor Cyan
$sharePointInfo = Get-SharePointInformation -Path $path -ServerName $ServerName

# Generate HTML Report
Write-Host "Generating HTML Report..." -ForegroundColor Cyan

# Pre-generate all required icons
Write-Host "Loading icons..." -ForegroundColor Yellow
$iconChart = Get-IconHTML -IconName 'chart'
$iconClipboard = Get-IconHTML -IconName 'clipboard'
$iconBuilding = Get-IconHTML -IconName 'building'
$iconComputer = Get-IconHTML -IconName 'computer'
$iconGear = Get-IconHTML -IconName 'gear'
$iconSearch = Get-IconHTML -IconName 'search'
$iconGlobe = Get-IconHTML -IconName 'globe'
$iconFolder = Get-IconHTML -IconName 'folder'
$iconUsers = Get-IconHTML -IconName 'users'
$iconUser = Get-IconHTML -IconName 'user'
$iconDatabase = Get-IconHTML -IconName 'database'
$iconFile = Get-IconHTML -IconName 'file'
$iconLock = Get-IconHTML -IconName 'lock'
$iconShield = Get-IconHTML -IconName 'shield'
$iconCrown = Get-IconHTML -IconName 'crown'
$iconPackage = Get-IconHTML -IconName 'package'
$iconPlug = Get-IconHTML -IconName 'plug'
$iconLightning = Get-IconHTML -IconName 'lightning'
$iconWrench = Get-IconHTML -IconName 'wrench'
$iconClock = Get-IconHTML -IconName 'clock'
$iconHospital = Get-IconHTML -IconName 'hospital'
$iconRefresh = Get-IconHTML -IconName 'refresh'
$iconTarget = Get-IconHTML -IconName 'target'
$iconFolder2 = Get-IconHTML -IconName 'folder2'
Write-Host "Icons loaded successfully" -ForegroundColor Green

$reportTitle = "SharePoint Health Check Report"

$fullHtml = @"
<!DOCTYPE html>
<html>
<head>
    <title>$reportTitle for SharePoint Server</title>    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }        
        /* Icon Styles */
        .icon-emoji {
            display: inline-block;
            vertical-align: middle;
            margin-right: 8px;
            font-size: 16px;
        }
        
        .nav-item .icon-emoji, .nav-submenu-item .icon-emoji {
            margin-right: 10px;
            font-size: 14px;
        }
        
        .section-header .icon-emoji, .subsection-header .icon-emoji {
            margin-right: 10px;
            font-size: 18px;
        }
        
        h2 .icon-emoji {
            font-size: 20px;
        }
        
        h3 .icon-emoji {
            font-size: 16px;
        }
        
        body {
            font-family: 'Segoe UI', sans-serif; 
            background-color: #f8f9fa; 
            line-height: 1.6;
        }
        
        /* Header Styles */
        .header {
            background: linear-gradient(135deg, #0078D7 0%, #005a9e 100%);
            color: white;
            padding: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .header h1 {
            margin-bottom: 10px;
            font-size: 2.5em;
        }
        
        .header-info {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }
        
        .header-info p {
            background: rgba(255, 255, 255, 0.1);
            padding: 8px 12px;
            border-radius: 5px;
            margin: 0;
        }
        
        /* Layout Container */
        .container {
            display: flex;
            min-height: calc(100vh - 140px);
        }
        
        /* Left Navigation Menu */
        .nav-menu {
            width: 320px;
            background: #2d3748;
            color: white;
            padding: 0;
            box-shadow: 2px 0 10px rgba(0,0,0,0.1);
            transition: width 0.3s ease;
            position: relative;
        }
        
        .nav-menu.collapsed {
            width: 60px;
        }
        
        .nav-toggle {
            background: #4299e1;
            color: white;
            border: none;
            padding: 15px;
            width: 100%;
            text-align: left;
            cursor: pointer;
            font-size: 16px;
            font-weight: bold;
            transition: background-color 0.3s;
        }
        
        .nav-toggle:hover {
            background: #3182ce;
        }
          .nav-toggle:after {
            content: '\25C0';
            float: right;
            transition: transform 0.3s;
        }
        
        .nav-menu.collapsed .nav-toggle:after {
            transform: rotate(180deg);
        }
        
        .nav-items {
            overflow: hidden;
            transition: all 0.3s ease;
        }
        
        .nav-menu.collapsed .nav-items {
            opacity: 0;
        }
          .nav-item {
            border-bottom: 1px solid #4a5568;
        }
        
        .nav-item a {
            display: block;
            padding: 15px 20px;
            color: #e2e8f0;
            text-decoration: none;
            transition: all 0.3s;
            cursor: pointer;
        }
        
        .nav-item a:hover {
            background: #4a5568;
            color: white;
            padding-left: 25px;
        }
        
        .nav-item.active a {
            background: #4299e1;
            color: white;
            border-left: 4px solid #63b3ed;
        }
          /* Sub-section styles */
        .nav-item.has-submenu > a:after {
            content: '\25B6';
            float: right;
            transition: transform 0.3s;
            font-size: 12px;
        }
        
        .nav-item.has-submenu.expanded > a:after {
            transform: rotate(90deg);
        }
        
        .nav-submenu {
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.3s ease;
            background: #1a202c;
        }
        
        .nav-item.expanded .nav-submenu {
            max-height: 500px;
        }
        
        .nav-submenu-item {
            border-bottom: 1px solid #2d3748;
        }
        
        .nav-submenu-item a {
            padding: 12px 40px;
            font-size: 14px;
            color: #cbd5e0;
            background: transparent;
        }
        
        .nav-submenu-item a:hover {
            background: #2d3748;
            color: white;
            padding-left: 45px;
        }
        
        .nav-submenu-item.active a {
            background: #2b77e6;
            color: white;
            border-left: 3px solid #63b3ed;
        }
        
        .section-count {
            float: right;
            background: #4299e1;
            color: white;
            padding: 2px 8px;
            border-radius: 10px;
            font-size: 12px;
            margin-top: 2px;
        }
        
        /* Main Content Area */
        .main-content {
            flex: 1;
            padding: 20px;
            background: white;
            margin: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        /* Summary Cards */
        .summary-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 15px;
            margin-bottom: 30px;
            box-shadow: 0 8px 25px rgba(0, 0, 0, 0.15);
        }
        
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        
        .summary-item {
            background: rgba(255, 255, 255, 0.15);
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            backdrop-filter: blur(10px);
        }
        
        .summary-number {
            font-size: 2.5em;
            font-weight: bold;
            display: block;
            margin-bottom: 5px;
        }
        
        .summary-label {
            font-size: 1em;
            opacity: 0.9;
        }
          /* Section Styles */
        .section {
            margin-bottom: 30px;
            display: none;
        }
        
        .section.active {
            display: block;
            animation: fadeIn 0.3s ease-in;
        }
        
        .subsection {
            margin-bottom: 25px;
            display: none;
        }
        
        .subsection.active {
            display: block;
            animation: fadeIn 0.3s ease-in;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .section-header {
            background: linear-gradient(135deg, #0078D7 0%, #005a9e 100%);
            color: white;
            padding: 20px;
            border-radius: 10px 10px 0 0;
            margin-bottom: 0;
        }
        
        .section-header h2 {
            margin: 0;
            font-size: 1.8em;
        }
        
        .subsection-header {
            background: linear-gradient(135deg, #4299e1 0%, #3182ce 100%);
            color: white;
            padding: 15px 20px;
            border-radius: 8px 8px 0 0;
            margin-bottom: 0;
            margin-top: 20px;
        }
        
        .subsection-header:first-child {
            margin-top: 0;
        }
        
        .subsection-header h3 {
            margin: 0;
            font-size: 1.3em;
        }
        
        .section-content {
            background: white;
            border: 1px solid #e2e8f0;
            border-top: none;
            border-radius: 0 0 10px 10px;
            padding: 20px;
        }
        
        .subsection-content {
            background: white;
            border: 1px solid #e2e8f0;
            border-top: none;
            border-radius: 0 0 8px 8px;
            padding: 15px 20px;
        }
        
        /* Table Styles */
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin: 15px 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }
        
        th, td { 
            padding: 12px 15px; 
            text-align: left; 
            border-bottom: 1px solid #e2e8f0;
        }
        
        th { 
            background: linear-gradient(135deg, #4299e1 0%, #3182ce 100%);
            color: white;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.9em;
        }
        
        tr:nth-child(even) { 
            background-color: #f8f9fa; 
        }
        
        tr:hover { 
            background-color: #e3f2fd; 
            transition: background-color 0.2s;
        }
        
        /* Responsive Design */
        @media (max-width: 768px) {
            .container {
                flex-direction: column;
            }
            
            .nav-menu {
                width: 100%;
                order: 2;
            }
            
            .main-content {
                order: 1;
                margin: 10px;
            }
            
            .summary-grid {
                grid-template-columns: 1fr;
            }
        }
        
        /* Footer */
        .footer {
            background: #2d3748;
            color: white;
            padding: 20px;
            text-align: center;
            margin-top: 30px;
        }
        
        .footer-info {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            max-width: 1200px;
            margin: 0 auto;
        }
        
        .footer-item {
            background: rgba(255, 255, 255, 0.1);
            padding: 15px;
            border-radius: 8px;
        }
        
        .footer-item strong {
            color: #63b3ed;
        }
    </style>    <script>
        // Navigation functionality
        document.addEventListener('DOMContentLoaded', function() {
            // Toggle navigation menu
            const navToggle = document.querySelector('.nav-toggle');
            const navMenu = document.querySelector('.nav-menu');
            
            navToggle.addEventListener('click', function() {
                navMenu.classList.toggle('collapsed');
            });
            
            // Handle main section and sub-section navigation
            const navLinks = document.querySelectorAll('.nav-item a, .nav-submenu-item a');
            const sections = document.querySelectorAll('.section');
            const subsections = document.querySelectorAll('.subsection');
            
            // Show summary section by default
            document.getElementById('summary').classList.add('active');
            document.querySelector('[data-section="summary"]').parentElement.classList.add('active');
            
            navLinks.forEach(link => {
                link.addEventListener('click', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    const sectionId = this.getAttribute('data-section');
                    const subsectionId = this.getAttribute('data-subsection');
                    
                    // Handle submenu toggle for main section links with submenus
                    if (this.parentElement.classList.contains('has-submenu') && !subsectionId) {
                        this.parentElement.classList.toggle('expanded');
                        return;
                    }
                    
                    // Remove active class from all nav items and sections
                    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
                    document.querySelectorAll('.nav-submenu-item').forEach(item => item.classList.remove('active'));
                    sections.forEach(s => s.classList.remove('active'));
                    subsections.forEach(s => s.classList.remove('active'));
                    
                    // Show the main section
                    const targetSection = document.getElementById(sectionId);
                    if (targetSection) {
                        targetSection.classList.add('active');
                    }
                    
                    // Handle subsection navigation
                    if (subsectionId) {
                        // Add active class to the submenu item
                        this.parentElement.classList.add('active');
                        
                        // Make sure the parent menu is expanded
                        const parentNavItem = this.closest('.nav-item.has-submenu');
                        if (parentNavItem) {
                            parentNavItem.classList.add('expanded');
                        }
                        
                        // Hide all subsections in the target section
                        const targetSectionSubsections = targetSection.querySelectorAll('.subsection');
                        targetSectionSubsections.forEach(sub => sub.classList.remove('active'));
                        
                        // Show the specific subsection
                        const targetSubsection = document.getElementById(subsectionId);
                        if (targetSubsection) {
                            targetSubsection.classList.add('active');
                        }
                    } else {
                        // If no subsection specified, show all subsections in the section
                        const targetSectionSubsections = targetSection.querySelectorAll('.subsection');
                        targetSectionSubsections.forEach(sub => sub.classList.add('active'));
                        
                        // Add active class to the main nav item
                        this.parentElement.classList.add('active');
                    }
                    
                    // Scroll to top of main content
                    document.querySelector('.main-content').scrollTop = 0;
                });
            });
            
            // Auto-expand first menu item with submenus on load
            const firstSubmenu = document.querySelector('.nav-item.has-submenu');
            if (firstSubmenu) {
                firstSubmenu.classList.add('expanded');
            }
        });
    </script>
</head>
<body>
    <div class="header">
        <h1>$reportTitle</h1>
        <div class="header-info">
            <p><strong>Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p><strong>Host Name:</strong> $(hostname)</p>
            <p><strong>User:</strong> $(whoami)</p>
            <p><strong>Scope:</strong> $reportScope</p>
        </div>
    </div>
    
    <div class="container">        <nav class="nav-menu">
            <button class="nav-toggle">$iconChart Navigation Menu</button>
            <div class="nav-items">
                <div class="nav-item">
                    <a href="#" data-section="summary">$iconClipboard Executive Summary</a>
                </div><div class="nav-item has-submenu">
                    <a href="#" data-section="farm-info">$iconBuilding SharePoint Farm Information <span class="section-count">3</span></a>
                    <div class="nav-submenu">
                        <div class="nav-submenu-item">
                            <a href="#" data-section="farm-info" data-subsection="SPServers">$iconComputer SharePoint Servers</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="farm-info" data-subsection="CentralAdmin">$iconGear Central Administration</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="farm-info" data-subsection="DiagnosticConfig">$iconSearch Diagnostic Configuration</a>
                        </div>
                    </div>
                </div>                <div class="nav-item has-submenu">
                    <a href="#" data-section="web-apps">$iconGlobe Web Applications & Sites <span class="section-count">4</span></a>
                    <div class="nav-submenu">
                        <div class="nav-submenu-item">
                            <a href="#" data-section="web-apps" data-subsection="WebApplications">$iconGlobe Web Applications</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="web-apps" data-subsection="SiteCollections">$iconFolder Site Collections</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="web-apps" data-subsection="SiteAdmins">$iconUsers Site Collection Admins</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="web-apps" data-subsection="SiteUsers">$iconUser Site Users</a>
                        </div>
                    </div>
                </div>                <div class="nav-item has-submenu">
                    <a href="#" data-section="databases">$iconDatabase Databases & Content <span class="section-count">3</span></a>
                    <div class="nav-submenu">
                        <div class="nav-submenu-item">
                            <a href="#" data-section="databases" data-subsection="SPDatabases">$iconDatabase Content Databases</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="databases" data-subsection="SPContentTypes">$iconFile Content Types</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="databases" data-subsection="SPBackupHistory">$iconDatabase Backup History</a>
                        </div>
                    </div>
                </div>                <div class="nav-item has-submenu">
                    <a href="#" data-section="security">$iconLock Security Configuration <span class="section-count">4</span></a>
                    <div class="nav-submenu">
                        <div class="nav-submenu-item">
                            <a href="#" data-section="security" data-subsection="SPSecurity">$iconShield Security Configuration</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="security" data-subsection="SPFarmAdmins">$iconCrown Farm Administrators</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="security" data-subsection="SPWebAppPolicies">$iconClipboard Web App Policies</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="security" data-subsection="TLSSettings">$iconLock TLS Settings</a>
                        </div>
                    </div>
                </div>                <div class="nav-item has-submenu">
                    <a href="#" data-section="services">$iconGear Services & Solutions <span class="section-count">5</span></a>
                    <div class="nav-submenu">
                        <div class="nav-submenu-item">
                            <a href="#" data-section="services" data-subsection="SPServiceAccounts">$iconUser Service Accounts</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="services" data-subsection="ServerServices">$iconGear Server Services</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="services" data-subsection="SPSolutions">$iconPackage SharePoint Solutions</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="services" data-subsection="SPFeatures">$iconPlug SharePoint Features</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="services" data-subsection="SPUserProfiles">$iconUsers User Profiles</a>
                        </div>
                    </div>
                </div>                <div class="nav-item has-submenu">
                    <a href="#" data-section="performance">$iconLightning Performance & Caching <span class="section-count">4</span></a>
                    <div class="nav-submenu">
                        <div class="nav-submenu-item">
                            <a href="#" data-section="performance" data-subsection="SQLCounters">$iconChart SQL Performance Counters</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="performance" data-subsection="SPCacheSettings">$iconLightning Cache Settings</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="performance" data-subsection="SPBlobCache">$iconDatabase BLOB Cache</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="performance" data-subsection="SPSearchTopology">$iconSearch Search Topology</a>
                        </div>
                    </div>
                </div>                <div class="nav-item has-submenu">
                    <a href="#" data-section="infrastructure">$iconWrench Infrastructure & Monitoring <span class="section-count">4</span></a>
                    <div class="nav-submenu">
                        <div class="nav-submenu-item">
                            <a href="#" data-section="infrastructure" data-subsection="WebBindings">$iconGlobe Web Bindings</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="infrastructure" data-subsection="SPIISSettings">$iconComputer IIS Settings</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="infrastructure" data-subsection="SPTimerJobs">$iconClock Timer Jobs</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="infrastructure" data-subsection="SPHealthAnalyzer">$iconHospital Health Analyzer</a>
                        </div>
                    </div>
                </div>                <div class="nav-item has-submenu">
                    <a href="#" data-section="patch-mgmt">$iconRefresh Patch Management <span class="section-count">2</span></a>
                    <div class="nav-submenu">                        <div class="nav-submenu-item">
                            <a href="#" data-section="patch-mgmt" data-subsection="FarmVersion">$iconBuilding Farm Version</a>
                        </div>
                        <div class="nav-submenu-item">
                            <a href="#" data-section="patch-mgmt" data-subsection="MissingUpdates">$iconRefresh Missing Windows Updates</a>
                        </div>
                    </div>
                </div>
            </div>
        </nav>
        
        <main class="main-content">
            <!-- Executive Summary Section -->
            <div id="summary" class="section">                <div class="summary-card">
                    <h2 style="color: white; margin-bottom: 20px;">$iconChart SharePoint Environment Summary</h2>
                    <div class="summary-grid">                        <div class="summary-item">
                            <span class="summary-number">31</span>
                            <span class="summary-label">Assessment Categories</span>
                        </div>
                        <div class="summary-item">
                            <span class="summary-number">32</span>
                            <span class="summary-label">CSV Reports Generated</span>
                        </div>
                        <div class="summary-item">
                            <span class="summary-number">1</span>
                            <span class="summary-label">HTML Report</span>
                        </div>
                    </div>
                </div>
                  <div class="section-header">
                    <h2>$iconClipboard Assessment Overview</h2>
                </div>
                <div class="section-content">
                    <p>This comprehensive SharePoint assessment report provides detailed analysis across 9 major categories covering 32 different assessment areas. The report includes both technical configuration details and security posture analysis.</p>
                    <br>
                    <h3>$iconFolder2 Report Components:</h3>
                    <ul style="margin-left: 20px; margin-top: 10px;">
                        <li><strong>Interactive HTML Report:</strong> This comprehensive report with collapsible navigation</li>
                        <li><strong>CSV Data Files:</strong> 33 detailed CSV files for data analysis and integration</li>
                        <li><strong>Assessment Categories:</strong> Covering Farm, Web Apps, Security, Performance, and more</li>
                    </ul>
                    <br>
                    <h3>$iconTarget Key Areas Assessed:</h3>
                    <ul style="margin-left: 20px; margin-top: 10px;">
                        <li>SharePoint Farm configuration and topology</li>
                        <li>Web applications and site collections</li>
                        <li>Security configuration and policies</li>
                        <li>Service accounts and permissions</li>
                        <li>Performance and caching settings</li>
                        <li>Infrastructure and monitoring</li>
                        <li>Patch management and updates</li>
                    </ul>
                </div>
            </div>
              <!-- SharePoint Farm Information Section -->            <div id="farm-info" class="section">
                <div class="section-header">
                    <h2>$iconBuilding SharePoint Farm Information</h2>
                </div>
                <div class="section-content">
                    <div id="SPServers" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconComputer SharePoint Servers</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPServers'])
                        </div>
                    </div>
                      <div id="CentralAdmin" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconGear Central Administration</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['CentralAdmin'])
                        </div>
                    </div>
                    
                    <div id="DiagnosticConfig" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconSearch Diagnostic Configuration</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['DiagnosticConfig'])
                        </div>
                    </div>
                </div>
            </div>
              <!-- Web Applications & Sites Section -->            <div id="web-apps" class="section">
                <div class="section-header">
                    <h2>$iconGlobe Web Applications & Sites</h2>
                </div>
                <div class="section-content">
                    <div id="WebApplications" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconGlobe Web Applications</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['WebApplications'])
                        </div>
                    </div>
                      <div id="SiteCollections" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconFolder Site Collections</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SiteCollections'])
                        </div>
                    </div>
                    
                    <div id="SiteAdmins" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconUsers Site Collection Admins</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SiteAdmins'])
                        </div>
                    </div>
                    
                    <div id="SiteUsers" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconUser Site Users</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SiteUsers'])
                        </div>
                    </div>
                </div>
            </div>
              <!-- Databases & Content Section -->
            <div id="databases" class="section">
                <div class="section-header">
                    <h2>$iconDatabase Databases & Content</h2>
                </div>
                <div class="section-content">
                    <div id="SPDatabases" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconDatabase Content Databases</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPDatabases'])
                        </div>
                    </div>
                    
                    <div id="SPContentTypes" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconFile Content Types</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPContentTypes'])
                        </div>
                    </div>
                    
                    <div id="SPBackupHistory" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconDatabase Backup History</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPBackupHistory'])
                        </div>
                    </div>
                </div>
            </div>
              <!-- Security Configuration Section -->
            <div id="security" class="section">
                <div class="section-header">
                    <h2>$iconLock Security Configuration</h2>
                </div>
                <div class="section-content">
                    <div id="SPSecurity" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconShield Security Configuration</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPSecurity'])
                        </div>
                    </div>
                    
                    <div id="SPFarmAdmins" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconCrown Farm Administrators</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPFarmAdmins'])
                        </div>
                    </div>
                    
                    <div id="SPWebAppPolicies" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconClipboard Web App Policies</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPWebAppPolicies'])
                        </div>
                    </div>
                    
                    <div id="TLSSettings" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconLock TLS Settings</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['TLSSettings'])
                        </div>
                    </div>
                </div>
            </div>              <!-- Services & Solutions Section -->
            <div id="services" class="section">
                <div class="section-header">
                    <h2>$iconGear Services & Solutions</h2>
                </div>                <div class="section-content">
                    <div id="SPServiceAccounts" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconUser Service Accounts</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPServiceAccounts'])
                        </div>
                    </div>
                    
                    <div id="ServerServices" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconGear Server Services</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['ServerServices'])
                        </div>
                    </div>
                    
                    <div id="SPSolutions" class="subsection">
                        <div class="subsection-header">
                            <h3>$iconPackage SharePoint Solutions</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPSolutions'])
                        </div>
                    </div>
                    
                    <div id="SPFeatures" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'plug') SharePoint Features</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPFeatures'])
                        </div>
                    </div>
                    
                    <div id="SPUserProfiles" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'users') User Profiles</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPUserProfiles'])
                        </div>
                    </div>
                </div>
            </div>
              <!-- Performance & Caching Section -->
            <div id="performance" class="section">
                <div class="section-header">
                    <h2>$(Get-IconSVG -IconName 'lightning') Performance & Caching</h2>
                </div>
                <div class="section-content">
                    <div id="SQLCounters" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'chart') SQL Performance Counters</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SQLCounters'])
                        </div>
                    </div>
                    
                    <div id="SPCacheSettings" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'lightning') Cache Settings</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPCacheSettings'])
                        </div>
                    </div>
                    
                    <div id="SPBlobCache" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'database') BLOB Cache</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPBlobCache'])
                        </div>
                    </div>
                    
                    <div id="SPSearchTopology" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'search') Search Topology</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPSearchTopology'])
                        </div>
                    </div>
                </div>
            </div>              <!-- Infrastructure & Monitoring Section -->
            <div id="infrastructure" class="section">
                <div class="section-header">
                    <h2>$(Get-IconSVG -IconName 'wrench') Infrastructure & Monitoring</h2>
                </div>
                <div class="section-content">
                    <div id="WebBindings" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'globe') Web Bindings</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['WebBindings'])
                        </div>
                    </div>
                    
                    <div id="SPIISSettings" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'computer') IIS Settings</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPIISSettings'])
                        </div>
                    </div>
                    
                    <div id="SPTimerJobs" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'clock') Timer Jobs</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPTimerJobs'])
                        </div>
                    </div>
                    
                    <div id="SPHealthAnalyzer" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'hospital') Health Analyzer</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['SPHealthAnalyzer'])
                        </div>
                    </div>
                </div>
            </div>              <!-- Patch Management Section -->
            <div id="patch-mgmt" class="section">
                <div class="section-header">
                    <h2>$(Get-IconSVG -IconName 'refresh') Patch Management</h2>
                </div>
                <div class="section-content">                    <div id="FarmVersion" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'building') Farm Version</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['FarmVersion'])
                        </div>
                    </div>
                    
                    <div id="MissingUpdates" class="subsection">
                        <div class="subsection-header">
                            <h3>$(Get-IconSVG -IconName 'refresh') Missing Windows Updates</h3>
                        </div>
                        <div class="subsection-content">
                            $($sharePointInfo['MissingUpdates'])
                        </div>
                    </div>
                </div>
            </div>
        </main>
    </div>
    
    <div class="footer">
        <div class="footer-info">
            <div class="footer-item">
                <strong>Report Generated:</strong><br>
                $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            </div>
            <div class="footer-item">
                <strong>SharePoint Server:</strong><br>
                $(hostname)
            </div>
            <div class="footer-item">
                <strong>Report Path:</strong><br>
                $htmlFile
            </div>
            <div class="footer-item">
                <strong>CSV Files Location:</strong><br>
                $path
            </div>
        </div>
    </div>
</body>
</html>
"@

# Write HTML file
try {
    $fullHtml | Out-File -FilePath $htmlFile -Encoding UTF8
    Write-Host "HTML report generated successfully!" -ForegroundColor Green
    Write-Host "Report saved to: $htmlFile" -ForegroundColor Yellow
} catch {
    Write-Host "Error generating HTML report: $($_.Exception.Message)" -ForegroundColor Red
}

# ----------------------------
# COMPLETION SUMMARY
# ----------------------------

Write-Host "`n=== SHAREPOINT ASSESSMENT COMPLETED ===" -ForegroundColor Green
Write-Host "Files generated:" -ForegroundColor Cyan
Write-Host "- HTML Report: $htmlFile" -ForegroundColor White
Write-Host "- CSV Files:" -ForegroundColor Cyan
Write-Host "  * SharePoint Servers: $csvSPServers" -ForegroundColor Gray
Write-Host "  * Central Admin: $csvCentralAdmin" -ForegroundColor Gray
Write-Host "  * Web Applications: $csvWebApps" -ForegroundColor Gray
Write-Host "  * Site Collections: $csvSiteCollections" -ForegroundColor Gray
Write-Host "  * Site Collection Admins: $csvSiteAdmins" -ForegroundColor Gray
Write-Host "  * Site Users: $csvSiteUsers" -ForegroundColor Gray
Write-Host "  * Farm Version: $csvFarmVersion" -ForegroundColor Gray
Write-Host "  * Service Accounts: $csvSPServiceAccounts" -ForegroundColor Gray
Write-Host "  * Server Services: $csvServerServices" -ForegroundColor Gray
Write-Host "  * Diagnostic Config: $csvDiagnosticConfig" -ForegroundColor Gray
Write-Host "  * Web Bindings: $csvWebBindings" -ForegroundColor Gray
Write-Host "  * TLS Settings: $csvTLSSettings" -ForegroundColor Gray
Write-Host "  * SQL Counters: $csvSQLCounters" -ForegroundColor Gray
Write-Host "  * SharePoint Solutions: $csvSPSolutions" -ForegroundColor Gray
Write-Host "  * SharePoint Features: $csvSPFeatures" -ForegroundColor Gray

Write-Host "- Security & Performance CSV Files:" -ForegroundColor Cyan
Write-Host "  * Content Databases: $csvSPDatabases" -ForegroundColor Gray
Write-Host "  * Security Configuration: $csvSPSecurity" -ForegroundColor Gray
Write-Host "  * Farm Administrators: $csvSPFarmAdmins" -ForegroundColor Gray
Write-Host "  * Web App Policies: $csvSPWebAppPolicies" -ForegroundColor Gray
Write-Host "  * Search Topology: $csvSPSearchTopology" -ForegroundColor Gray
Write-Host "  * Cache Settings: $csvSPCacheSettings" -ForegroundColor Gray
Write-Host "  * Timer Jobs: $csvSPTimerJobs" -ForegroundColor Gray
Write-Host "  * Health Analyzer: $csvSPHealthAnalyzer" -ForegroundColor Gray
Write-Host "  * Content Types: $csvSPContentTypes" -ForegroundColor Gray
Write-Host "  * IIS Settings: $csvSPIISSettings" -ForegroundColor Gray
Write-Host "  * Missing Updates: $csvMissingUpdates" -ForegroundColor Gray
Write-Host "  * User Profiles: $csvSPUserProfiles" -ForegroundColor Gray
Write-Host "  * BLOB Cache: $csvSPBlobCache" -ForegroundColor Gray

Write-Host "`nNext Steps:" -ForegroundColor Yellow
Write-Host "1. Review the HTML report: $htmlFile" -ForegroundColor White
Write-Host "2. Analyze individual CSV files for detailed data" -ForegroundColor White
Write-Host "3. Store reports securely and follow data retention policies" -ForegroundColor White
Write-Host "4. Schedule regular SharePoint health checks for ongoing monitoring" -ForegroundColor White

Write-Host "`nThank you for using the SharePoint Assessment Tool!" -ForegroundColor Green
Write-Host "Script created for SharePoint Health Assessment - Version 1.0" -ForegroundColor Gray
