######## POWER BI ########
##### External Tools #####
########## Code ##########
##########################

param (
    [string]$servername # Allow Power BI Desktop to pass the server name as a parameter
)

# Since you have already execute the get_started_checklist.ps1 
# You have already installed the SqlServer module and located the 2 required DLLs

# Load necessary assemblies for Analysis Services
# Update the DLL paths 

$core_dll_path = "C:\xxxx\MICROSOFT.ANALYSISSERVICES.CORE.DLL"
$tabular_dll_path = "C:\xxxx\MICROSOFT.ANALYSISSERVICES.TABULAR.DLL"

# Check if all dependencies are present:
if (-not(Test-Path -Path $core_dll_path)) {
    
    Write-Host "Error loading the MICROSOFT.ANALYSISSERVICES.CORE.DLL might be missing dependencies or the wrong version." -ForegroundColor Red
    
    Write-Host "Press Enter to close this window..." -ForegroundColor Green
    # Wait for user input to keep the window open
    Read-Host
}

if (-not(Test-Path -Path $tabular_dll_path)) {
    
    Write-Host "Error loading the MICROSOFT.ANALYSISSERVICES.TABULAR.DLL might be missing dependencies or the wrong version." -ForegroundColor Red
    
    Write-Host "Press Enter to close this window..." -ForegroundColor Green
    # Wait for user input to keep the window open
    Read-Host
}

Add-Type -Path $core_dll_path
Add-Type -Path $tabular_dll_path

# Connect to the Analysis Services instance
$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$connectionString = "DataSource=$serverName"

try {
    $server.Connect($connectionString)
    Write-Host "Connected to Power BI Analysis Services instance on $serverName" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to the Analysis Services instance: $_"
    return
}

# Extract measures and check for errors
$database = $server.Databases[0]  # Assuming only one database
$model = $database.Model

# List to store measures with errors
$measuresWithErrors = @()

foreach ($table in $model.Tables) {
    foreach ($measure in $table.Measures) {
        if ($measure.ErrorMessage -ne '') {
            # Add the measure with its error details to the list
            $errorDetails = @{
                "Measure Name"  = $measure.Name
                "Table Name"    = $table.Name
                "Folder Name"   = $measure.DisplayFolder
                "Error Message" = $measure.ErrorMessage
            }
            $measuresWithErrors += New-Object PSObject -Property $errorDetails
        }
    }
}

# Disconnect from the server
$server.Disconnect()
Write-Host "Disconnected from Power BI Analysis Services instance." -ForegroundColor Cyan

# Display measures with errors 
if ($measuresWithErrors.Count -gt 0) {
    Write-Host "`nMeasures with Errors:" -ForegroundColor Black -BackgroundColor Yellow
    $measuresWithErrors | Format-Table "Measure Name", "Table Name", "Folder Name", "Error Message" -AutoSize
    
} else {
    Write-Host "No Objects with errors found!" -ForegroundColor Black -BackgroundColor Yellow

}
    Write-Host "Press Enter to close this window..." -ForegroundColor Green

# Wait for user input to keep the window open
Read-Host