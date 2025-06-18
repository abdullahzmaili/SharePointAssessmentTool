# SharePoint Assessment Tool v1.0

A comprehensive PowerShell-based assessment tool for analyzing SharePoint Server environments. This tool provides detailed health checks, security assessments, and performance analysis across all major SharePoint components.

## 🚀 Features

### Comprehensive Assessment Coverage
- **SharePoint Farm Information** - Server topology, Central Admin configuration, diagnostic settings
- **Web Applications & Sites** - Complete analysis of web apps, site collections, administrators, and users
- **Security Configuration** - Authentication settings, farm administrators, web app policies, TLS configuration
- **Database Analysis** - Content databases, backup history, content types summary
- **Service Management** - Service accounts, server services, SharePoint solutions and features
- **Performance Monitoring** - SQL performance counters, cache settings, BLOB cache configuration
- **Infrastructure Details** - IIS settings, web bindings, timer jobs, health analyzer rules
- **Patch Management** - Farm version information and missing Windows updates analysis

### Output Formats
- **Interactive HTML Report** - Professional, responsive report with collapsible navigation
- **Individual CSV Files** - 32+ detailed CSV files for data analysis and integration
- **Modular Structure** - Easy to customize and extend assessment categories

### Security Features
- TLS/SSL configuration analysis
- Authentication provider assessment
- Service account enumeration
- Farm administrator identification
- Web application security policies
- File type restrictions analysis

## 📋 Prerequisites

### System Requirements
- **Operating System**: Windows Server 2012 R2 or later
- **PowerShell**: Version 5.1 or later
- **SharePoint**: SharePoint Server 2013, 2016, 2019, or Subscription Edition
- **Permissions**: Local Administrator and SharePoint Farm Administrator rights

### Required Modules
The script automatically attempts to load required modules:
- Microsoft.SharePoint.PowerShell (SharePoint Management Shell)
- SqlServer (for database queries)
- WebAdministration (for IIS configuration)

### Account Requirements
- **SharePoint Farm Administrator** privileges required
- **Local Administrator** rights on the SharePoint server
- **SQL Server** access for database-related assessments

## 🎯 Quick Start

1. **Download and Place Script**
   ```powershell
   # Place the script in your desired location
   # Example: C:\Scripts\SharePointAssessmentTool_V1.0.ps1
   ```

2. **Open SharePoint Management Shell as Administrator**
   ```powershell
   # Run SharePoint Management Shell as Administrator
   ```

3. **Execute the Script**
   ```powershell
   .\SharePointAssessmentTool_V1.0.ps1
   ```

4. **Specify Output Directory**
   ```
   Enter the full path (without filename) to save the reports (e.g., C:\temp): C:\SharePointReports
   ```

5. **Review Generated Reports**
   - Main HTML Report: `[ServerName]-SharePointReport.html`
   - Individual CSV files: `[ServerName]-[Category].csv`

## 📁 Output Structure

The tool generates the following files in your specified directory:

### HTML Report
- `[ServerName]-SharePointReport.html` - Interactive HTML report with all assessment data

### CSV Data Files (32 files)
- `[ServerName]-SPServers.csv` - SharePoint server information
- `[ServerName]-WebApplications.csv` - Web application details
- `[ServerName]-SiteCollections.csv` - Site collection inventory
- `[ServerName]-SPDatabases.csv` - Content database information
- `[ServerName]-SPSecurity.csv` - Security configuration details
- `[ServerName]-TLSSettings.csv` - TLS/SSL configuration
- `[ServerName]-SQLCounters.csv` - SQL performance metrics
- `[ServerName]-SPSolutions.csv` - SharePoint solutions inventory
- `[ServerName]-SPFeatures.csv` - SharePoint features status
- `[ServerName]-MissingUpdates.csv` - Missing Windows updates
- Plus 22 additional specialized CSV files covering all assessment areas

## 🛡️ Security Considerations

### Data Sensitivity
- Reports may contain sensitive configuration information
- Service account details are included in outputs
- Database connection strings and server details are captured
- Store reports in secure locations with appropriate access controls

### Network Security
- Tool performs local assessments only
- No data transmitted outside the assessed environment
- All queries are read-only operations

### Permissions Impact
- Requires elevated privileges for complete assessment
- Some checks may fail with insufficient permissions
- Error handling prevents script interruption

## 🔧 Customization

### Adding Custom Assessments
The modular design allows easy extension:
```powershell
# Add new assessment section
$customAssessment = Get-CustomSharePointData
$csvCustom = Join-Path -Path $Path -ChildPath "$ServerName-Custom.csv"
$customAssessment | Export-Csv -Path $csvCustom -NoTypeInformation
$htmlSections['Custom'] = $customAssessment | ConvertTo-Html -Fragment -PreContent "<h2>Custom Assessment</h2>"
```

### Modifying Output Paths
Edit the file path definitions section:
```powershell
$csvCustomFile = Join-Path -Path $path -ChildPath "$ServerName-CustomName.csv"
```

### Customizing HTML Styling
The HTML template includes comprehensive CSS that can be modified for corporate branding or styling preferences.

## 🚨 Troubleshooting

### Common Issues

**SharePoint PowerShell Module Not Found**
```powershell
# Ensure SharePoint Management Shell is installed
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
```

**Access Denied Errors**
- Verify Farm Administrator privileges
- Check local Administrator rights
- Ensure SQL Server connectivity

**Performance Issues**
- Large environments may take 15-30 minutes to complete
- Consider running during maintenance windows
- Monitor server resources during execution

**Incomplete Data Collection**
- Review PowerShell execution policy
- Check Windows Update service status
- Verify WMI service availability

### Error Handling
The script includes comprehensive error handling:
- Individual section failures don't stop overall assessment
- Error messages are captured in the HTML report
- Partial data collection continues even with some failures

## 📊 Report Navigation

### HTML Report Features
- **Collapsible Navigation Menu** - Left-side menu with expandable sections
- **Responsive Design** - Works on desktop and mobile devices
- **Search-Friendly Data** - All data in structured HTML tables
- **Professional Styling** - Clean, corporate-ready presentation

### Navigation Structure
- Executive Summary
- SharePoint Farm Information (3 subsections)
- Web Applications & Sites (4 subsections)
- Databases & Content (3 subsections)
- Security Configuration (4 subsections)
- Services & Solutions (5 subsections)
- Performance & Caching (4 subsections)
- Infrastructure & Monitoring (4 subsections)
- Patch Management (2 subsections)

## 📝 Version History

### v1.0 (Current)
- Initial release with comprehensive SharePoint assessment
- 32 assessment categories across 9 major sections
- Interactive HTML reporting with collapsible navigation
- Complete CSV data export capability
- Enhanced error handling and reporting
- Professional UI with responsive design

## 🤝 Support

### Documentation
- See `Instructions.md` for detailed usage instructions
- See `QuickStart.md` for rapid deployment guide

### Best Practices
- Run during maintenance windows for large environments
- Review all outputs for sensitive information before sharing
- Store reports securely with appropriate access controls
- Schedule regular assessments to track environment changes

### Contributing
The tool is designed for easy extension and customization. When adding new assessment categories:
1. Follow the existing pattern for data collection
2. Include appropriate error handling
3. Add CSV export capability
4. Update HTML navigation structure
5. Document new features

---

**Note**: This tool is designed for SharePoint on-premises environments. For SharePoint Online assessments, different approaches and tools are required.
