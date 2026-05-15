try {
    # 1. Get the folder where THIS script is saved
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $envPath = Join-Path $scriptDir ".env"

    # 2. Load and Clean .env file
    if (Test-Path $envPath) {
        Get-Content $envPath | Where-Object { $_ -match "=" } | ForEach-Object {
            $name, $value = $_.Split('=', 2)
            $cleanValue = $value.Trim().Replace("'", "").Replace('"', "")
            Set-Variable -Name "ENV_$($name.Trim())" -Value $cleanValue -ErrorAction SilentlyContinue
        }
    } else { 
        throw "ERROR: .env file not found at: $envPath"
    }

    # 3. Assign Source Path from .env
    $sourcePath = $ENV_source_a_path
    if (-not $sourcePath) { throw "source_a_path is missing in your .env file" }

    # 4. Define Subfolder Paths
    $searchablePath = Join-Path $sourcePath "01_processed_text_extractable"
    $archivePath = Join-Path $sourcePath "original_files_archive"

    # 5. Create folders if they don't exist
    if (-not (Test-Path $searchablePath)) { [void](New-Item -ItemType Directory -Path $searchablePath -Force) }
    if (-not (Test-Path $archivePath)) { [void](New-Item -ItemType Directory -Path $archivePath -Force) }

    # 6. Rename files without extensions to .pdf
    Get-ChildItem -Path $sourcePath | Where-Object { -not $_.PSIsContainer -and $_.Extension -eq "" } | ForEach-Object {
        Rename-Item -Path $_.FullName -NewName ($_.Name + ".pdf")
    }

    # 7. Activate Virtual Environment
    $venvPath = Join-Path $scriptDir ".venv\Scripts\Activate.ps1"
    if (Test-Path $venvPath) {
        & $venvPath
    }

    # 8. Process Files
    $files = Get-ChildItem -Path $sourcePath -Filter *.pdf | Where-Object { -not $_.PSIsContainer }
    
    if ($files.Count -eq 0) { Write-Host "No files found to process." -ForegroundColor Cyan }

    foreach ($file in $files) {
        $outputFull = Join-Path $searchablePath ("OCR_" + $file.BaseName + ".pdf")
        Write-Host "Processing: $($file.Name)..." -ForegroundColor Yellow
        
        # Run OCR
#	ocrmypdf --force-ocr --clean --deskew --tesseract-oem 1 "$($file.FullName)" "$outputFull"
	ocrmypdf --force-ocr --deskew --user-words "$scriptDir\words.txt" --tesseract-oem 1 --tesseract-pagesegmode 6 "$($file.FullName)" "$outputFull"


        # Move original only if OCR was successful
        if ($LASTEXITCODE -eq 0) {
            Move-Item -Path $file.FullName -Destination $archivePath -Force
            Write-Host "Success! Moved to archive." -ForegroundColor Green
        } else {
            Write-Host "Failed to process $($file.Name)" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Workflow Complete." -ForegroundColor Cyan
pause