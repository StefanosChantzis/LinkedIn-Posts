######## POWER BI ########
##### External Tools #####
####### Check List #######
##########################

Write-Host "SqlServer module - Check" -BackgroundColor White -ForegroundColor black
Start-Sleep -Seconds 2

## Check if the SqlServer module is installed
##########################
if (-not (Get-Module -ListAvailable -Name SqlServer)) {
    Write-Host "The 'SqlServer' module is not installed. Installing now..." -ForegroundColor Yellow

    try {
        # Install the SqlServer module
        Install-Module -Name SqlServer -Scope CurrentUser -Force -AllowClobber
        Write-Host "The 'SqlServer' module has been successfully installed." -ForegroundColor Green
    } catch {
        Write-Host "Failed to install the 'SqlServer' module. Error: $_" -ForegroundColor Red
    }
} else {
    Write-Host "The 'SqlServer' module is already installed." -ForegroundColor Green
}

Write-Host "DLLs - Check" -BackgroundColor White -ForegroundColor black
Start-Sleep -Seconds 2

## Check for necessary DLLs 
##########################
$filesToSearch = @("Microsoft.AnalysisServices.Core.dll", "Microsoft.AnalysisServices.Tabular.dll")

# Define the starting directory to search (default to C:\ for a full system search)
$searchDirectory = @("C:\Program Files (x86)", "C:\Program Files")

# Search for the files
Write-Host "Searching for the specified DLLs. This may take some time..." -ForegroundColor Yellow
try {
    foreach ($file in $filesToSearch) {
        $result = Get-ChildItem -Path $searchDirectory -Recurse -ErrorAction SilentlyContinue -Filter $file
        if ($result) {
            Write-Host "$file found at the following locations:" -ForegroundColor Green
            $result.FullName | ForEach-Object { Write-Host $_ }
        } else {
            Write-Host "$file not found on this machine. " -ForegroundColor Red
            Write-Host "Suggestion: Download and Install Microsoft SQL server Developer edition (free)" -BackgroundColor Yellow -ForegroundColor Black
            Write-Host "https://www.microsoft.com/en-us/sql-server/sql-server-downloads?msockid=1cddd30b3ae762fb0677c7ad3be7632e" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "An error occurred during the search: $_" -ForegroundColor Red
}

##########################
