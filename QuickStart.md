# SharePoint Assessment Tool - Quick Start Guide

Get up and running with the SharePoint Assessment Tool in under 5 minutes.

## ⚡ Rapid Deployment

### Step 1: Prerequisites Check
```powershell
# Verify you have required permissions
Get-SPFarm | Select-Object CurrentUserIsAdmin
# Should return: CurrentUserIsAdmin : True
```

### Step 2: Download and Setup
```powershell
# Navigate to your scripts directory
cd D:\Scripts\SharePointServerAssessmentTool

# Verify script is present
Get-ChildItem SharePointAssessmentTool_V1.0.ps1
```

### Step 3: Execute Assessment
```powershell
# Run the script (as Farm Administrator)
.\SharePointAssessmentTool_V1.0.ps1
```

### Step 4: Specify Output Location
```
Enter the full path (without filename) to save the reports (e.g., C:\temp): C:\SharePointReports
```

### Step 5: Wait for Completion
- **Small Environments**: 5-10 minutes
- **Medium Environments**: 10-20 minutes  
- **Large Environments**: 20-30 minutes

### Step 6: Review Results
```powershell
# Navigate to output directory
cd C:\SharePointReports

# View generated files
Get-ChildItem *.html, *.csv | Select-Object Name, Length, LastWriteTime
```

## 🎯 What You Get

### Main Report
- **HTML Report**: `[ServerName]-SharePointReport.html`
  - Interactive navigation
  - Professional presentation
  - Complete assessment data

### Data Files (32 CSV files)
- Server configuration details
- Security settings analysis  
- Performance metrics
- Database information
- Service configurations
- Patch management status

## 🚀 Immediate Actions

### 1. Open HTML Report
```powershell
# Open main report in default browser
Start-Process "C:\SharePointReports\[ServerName]-SharePointReport.html"
```

### 2. Review Critical Areas
**Priority 1: Security Configuration**
- Navigate to "Security Configuration" section
- Review Farm Administrators
- Check TLS Settings
- Verify Web App Policies

**Priority 2: Patch Management**
- Check "Patch Management" section
- Review Farm Version
- Identify Missing Updates

**Priority 3: Performance Issues**
- Check "Performance & Caching" section
- Review SQL Performance Counters
- Verify Cache Settings

### 3. Export Key Data
```powershell
# Copy critical CSV files for analysis
$criticalFiles = @(
    "*-SPSecurity.csv",
    "*-MissingUpdates.csv", 
    "*-SQLCounters.csv",
    "*-SPDatabases.csv"
)

foreach ($pattern in $criticalFiles) {
    Copy-Item $pattern C:\CriticalReports\
}
```

## 🔧 Quick Customization

### Change Output Location
Edit line 23 in the script:
```powershell
# Original
$path = Read-Host "Enter the full path (without filename) to save the reports (e.g., C:\temp)"

# Modified for automatic path
$path = "C:\SharePointReports\$(Get-Date -Format 'yyyy-MM-dd')"
```

### Run Silently
Add these lines at the beginning after line 20:
```powershell
# Silent execution with predefined path
$path = "C:\SharePointReports"
# Comment out the Read-Host line
```

### Schedule Regular Runs
```powershell
# Create scheduled task for weekly assessments
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File D:\Scripts\SharePointServerAssessmentTool\SharePointAssessmentTool_V1.0.ps1"
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2AM
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName "SharePoint Weekly Assessment" -Action $Action -Trigger $Trigger -Settings $Settings -User "DOMAIN\SPFarmAdmin" -RunLevel Highest
```

## 🚨 Quick Troubleshooting

### Script Won't Start
```powershell
# Check execution policy
Get-ExecutionPolicy

# If restricted, temporarily allow
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Permission Errors
```powershell
# Verify farm admin rights
Get-SPFarm | Select-Object CurrentUserIsAdmin

# Check if running as administrator
([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
```

### Missing Modules
```powershell
# Manually load SharePoint module
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Check if loaded successfully
Get-PSSnapin | Where-Object {$_.Name -like "*SharePoint*"}
```

### Incomplete Data
- **Some sections show errors**: Normal for environments with restricted access
- **Missing Windows Updates section fails**: Requires Windows Update service running
- **SQL counters unavailable**: SQL Server may not be local or accessible

## 📊 Quick Report Review

### 1. Executive Summary
- Shows 31 assessment categories
- 32 CSV reports generated
- Overall environment health snapshot

### 2. Priority Review Areas
**Immediate Attention Required:**
- Any "Error collecting" messages in red
- Missing critical updates in Patch Management
- Security misconfigurations
- Performance counter warnings

**Plan for Action:**
- Review all service accounts
- Plan update installation schedule
- Address security findings
- Optimize performance bottlenecks

### 3. Data Export for Further Analysis
```powershell
# Quick PowerShell analysis of key metrics
Import-Csv "*-SPDatabases.csv" | Where-Object {[int]$_.SizeGB -gt 50} | Select-Object Name, SizeGB, SiteCount

Import-Csv "*-MissingUpdates.csv" | Where-Object {$_.Severity -eq "Critical"} | Select-Object Title, KB

Import-Csv "*-SPSecurity.csv" | Where-Object {$_.AnonymousAccess -eq "True"} | Select-Object WebApplication, Zone
```

## 🎯 Next Steps

### 1. Full Review
- Schedule time to review complete HTML report
- Document findings and recommendations
- Plan remediation activities

### 2. Regular Assessments
- Set up automated monthly runs
- Track changes over time
- Monitor security posture

### 3. Advanced Usage
- See `Instructions.md` for detailed customization options
- See `README.md` for complete feature documentation

---

**Time to Complete**: ~5 minutes setup + assessment runtime (5-30 minutes depending on environment size)

**Output**: 1 HTML report + 32 CSV data files + Complete SharePoint assessment

**Next Action**: Open the HTML report and review the Executive Summary section!
