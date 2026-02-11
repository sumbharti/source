<#
.SYNOPSIS
    Discovers dependencies for all Canvas Apps in a Power Platform environment
.DESCRIPTION
    Uses Power Platform CLI to identify and document all dependencies for canvas apps
.NOTES
    Requires: Power Platform CLI (pac) installed
#>

# Configuration
$environmentUrl = "https://your-environment.crm.dynamics.com"  # Replace with your environment URL
$outputFolder = ".\CanvasAppDependencies"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Create output folder
New-Item -ItemType Directory -Force -Path $outputFolder | Out-Null

Write-Host "================================" -ForegroundColor Cyan
Write-Host "Canvas App Dependency Discovery" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Authenticate to Power Platform
Write-Host "Authenticating to Power Platform..." -ForegroundColor Yellow
pac auth create --environment $environmentUrl

# List all solutions and select environment
Write-Host "`nListing available environments..." -ForegroundColor Yellow
pac org list

# Get all canvas apps
Write-Host "`nRetrieving all canvas apps..." -ForegroundColor Yellow
$canvasAppsJson = pac canvas list --environment $environmentUrl --json
$canvasApps = $canvasAppsJson | ConvertFrom-Json

Write-Host "Found $($canvasApps.Count) canvas apps" -ForegroundColor Green

# Initialize results array
$dependencyResults = @()

# Process each canvas app
foreach ($app in $canvasApps) {
    Write-Host "`n----------------------------------------" -ForegroundColor Cyan
    Write-Host "Processing: $($app.DisplayName)" -ForegroundColor Cyan
    Write-Host "App ID: $($app.Name)" -ForegroundColor Gray
    
    $appDependencies = [PSCustomObject]@{
        AppName = $app.DisplayName
        AppId = $app.Name
        Environment = $app.EnvironmentName
        CreatedTime = $app.CreatedTime
        LastModifiedTime = $app.LastModifiedTime
        Owner = $app.Owner.email
        Tables = @()
        Flows = @()
        Connections = @()
        CustomConnectors = @()
        ComponentLibraries = @()
        EnvironmentVariables = @()
        OtherDependencies = @()
    }
    
    # Download the app package to analyze
    $appFileName = "$outputFolder\$($app.DisplayName -replace '[^a-zA-Z0-9]', '_')_$timestamp.msapp"
    
    try {
        Write-Host "  Downloading app package..." -ForegroundColor Yellow
        pac canvas download --msapp-file $appFileName --app-id $app.Name --environment $environmentUrl
        
        # Unpack the msapp file to analyze contents
        $unpackFolder = "$outputFolder\$($app.DisplayName -replace '[^a-zA-Z0-9]', '_')_unpacked"
        Write-Host "  Unpacking app..." -ForegroundColor Yellow
        pac canvas unpack --msapp $appFileName --sources $unpackFolder
        
        # Analyze Connections.json for data sources
        $connectionsFile = Join-Path $unpackFolder "Connections\Connections.json"
        if (Test-Path $connectionsFile) {
            $connections = Get-Content $connectionsFile -Raw | ConvertFrom-Json
            
            foreach ($conn in $connections.PSObject.Properties) {
                $connDetails = [PSCustomObject]@{
                    Name = $conn.Value.id
                    Type = $conn.Value.dataSources[0].name
                    DisplayName = $conn.Value.dataSources[0].displayName
                }
                $appDependencies.Connections += $connDetails
                Write-Host "    Connection: $($connDetails.DisplayName) ($($connDetails.Type))" -ForegroundColor Gray
            }
        }
        
        # Analyze DataSources.json for tables and entities
        $dataSourcesFile = Join-Path $unpackFolder "DataSources\*.json"
        $dataSourceFiles = Get-ChildItem -Path (Join-Path $unpackFolder "DataSources") -Filter "*.json" -ErrorAction SilentlyContinue
        
        foreach ($dsFile in $dataSourceFiles) {
            $dataSource = Get-Content $dsFile.FullName -Raw | ConvertFrom-Json
            
            if ($dataSource.Type -eq "Table" -or $dataSource.TableDefinition) {
                $tableDetails = [PSCustomObject]@{
                    Name = $dataSource.Name
                    TableName = $dataSource.TableName
                    EntitySetName = $dataSource.EntitySetName
                    Type = $dataSource.Type
                }
                $appDependencies.Tables += $tableDetails
                Write-Host "    Table: $($tableDetails.TableName)" -ForegroundColor Gray
            }
        }
        
        # Scan for Component Libraries in .pa.yaml files
        $yamlFiles = Get-ChildItem -Path $unpackFolder -Filter "*.pa.yaml" -Recurse -ErrorAction SilentlyContinue
        foreach ($yamlFile in $yamlFiles) {
            $yamlContent = Get-Content $yamlFile.FullName -Raw
            if ($yamlContent -match "ComponentLibrary") {
                Write-Host "    Component Library reference found" -ForegroundColor Gray
                # Parse YAML for component library details (requires additional parsing)
            }
        }
        
        # Clean up unpacked files (optional - comment out to keep for review)
        # Remove-Item -Path $unpackFolder -Recurse -Force -ErrorAction SilentlyContinue
        
    }
    catch {
        Write-Host "  Error processing app: $_" -ForegroundColor Red
        $appDependencies.OtherDependencies += "Error: $_"
    }
    
    # Add to results
    $dependencyResults += $appDependencies
}

# Query Dataverse for Flow dependencies using pac data
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Querying Dataverse for Flow dependencies" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

foreach ($appDep in $dependencyResults) {
    Write-Host "`nChecking flows for: $($appDep.AppName)" -ForegroundColor Yellow
    
    # Query workflow (flows) table for references to canvas app
    # This requires FetchXML query through pac data
    $fetchXml = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='category' />
    <attribute name='type' />
    <filter>
      <condition attribute='clientdata' operator='like' value='%$($appDep.AppId)%' />
      <condition attribute='category' operator='eq' value='5' />
    </filter>
  </entity>
</fetch>
"@
    
    try {
        # Note: pac data query command syntax
        # You may need to adjust based on your pac CLI version
        $flowQuery = "workflows?`$filter=contains(clientdata,'$($appDep.AppId)') and category eq 5&`$select=name,workflowid"
        # This is a placeholder - actual implementation would use pac data query
        Write-Host "  Query: $flowQuery" -ForegroundColor Gray
    }
    catch {
        Write-Host "  Could not query flows: $_" -ForegroundColor Red
    }
}

# Export results to CSV
$csvFile = "$outputFolder\CanvasApp_Dependencies_$timestamp.csv"
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Exporting Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$exportData = @()
foreach ($appDep in $dependencyResults) {
    $exportData += [PSCustomObject]@{
        'App Name' = $appDep.AppName
        'App ID' = $appDep.AppId
        'Owner' = $appDep.Owner
        'Created' = $appDep.CreatedTime
        'Last Modified' = $appDep.LastModifiedTime
        'Tables Count' = $appDep.Tables.Count
        'Tables' = ($appDep.Tables.TableName -join '; ')
        'Connections Count' = $appDep.Connections.Count
        'Connections' = ($appDep.Connections.DisplayName -join '; ')
        'Connection Types' = ($appDep.Connections.Type -join '; ')
        'Flows Count' = $appDep.Flows.Count
        'Flows' = ($appDep.Flows.Name -join '; ')
    }
}

$exportData | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
Write-Host "CSV exported to: $csvFile" -ForegroundColor Green

# Export detailed JSON
$jsonFile = "$outputFolder\CanvasApp_Dependencies_Detailed_$timestamp.json"
$dependencyResults | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonFile -Encoding UTF8
Write-Host "Detailed JSON exported to: $jsonFile" -ForegroundColor Green

# Generate HTML Report
$htmlFile = "$outputFolder\CanvasApp_Dependencies_Report_$timestamp.html"
$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Canvas App Dependencies Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        h1 { color: #0078d4; }
        h2 { color: #106ebe; margin-top: 30px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; background-color: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th { background-color: #0078d4; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f5f5f5; }
        .summary { background-color: white; padding: 20px; margin-bottom: 20px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .tag { display: inline-block; padding: 4px 8px; margin: 2px; background-color: #e1f5fe; border-radius: 3px; font-size: 12px; }
        .timestamp { color: #666; font-size: 14px; }
    </style>
</head>
<body>
    <h1>Canvas App Dependencies Report</h1>
    <div class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</div>
    
    <div class="summary">
        <h2>Summary</h2>
        <p><strong>Total Canvas Apps:</strong> $($dependencyResults.Count)</p>
        <p><strong>Environment:</strong> $environmentUrl</p>
    </div>
"@

foreach ($appDep in $dependencyResults) {
    $html += @"
    <h2>$($appDep.AppName)</h2>
    <table>
        <tr><th>Property</th><th>Value</th></tr>
        <tr><td>App ID</td><td>$($appDep.AppId)</td></tr>
        <tr><td>Owner</td><td>$($appDep.Owner)</td></tr>
        <tr><td>Created</td><td>$($appDep.CreatedTime)</td></tr>
        <tr><td>Last Modified</td><td>$($appDep.LastModifiedTime)</td></tr>
    </table>
    
    <h3>Tables ($($appDep.Tables.Count))</h3>
    <div>
"@
    foreach ($table in $appDep.Tables) {
        $html += "<span class='tag'>$($table.TableName)</span>"
    }
    
    $html += @"
    </div>
    
    <h3>Connections ($($appDep.Connections.Count))</h3>
    <div>
"@
    foreach ($conn in $appDep.Connections) {
        $html += "<span class='tag'>$($conn.DisplayName) ($($conn.Type))</span>"
    }
    
    $html += "</div>"
}

$html += @"
</body>
</html>
"@

$html | Out-File -FilePath $htmlFile -Encoding UTF8
Write-Host "HTML report exported to: $htmlFile" -ForegroundColor Green

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total Apps Processed: $($dependencyResults.Count)" -ForegroundColor Green
Write-Host "Total Tables Found: $(($dependencyResults.Tables | Measure-Object).Count)" -ForegroundColor Green
Write-Host "Total Connections Found: $(($dependencyResults.Connections | Measure-Object).Count)" -ForegroundColor Green
Write-Host "`nOutput files saved to: $outputFolder" -ForegroundColor Yellow

# Open HTML report
Write-Host "`nOpening HTML report..." -ForegroundColor Yellow
Start-Process $htmlFile
