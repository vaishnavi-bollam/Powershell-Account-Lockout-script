# Specify the path to the input CSV file
$inputCsvPath = "C:\Users\Cablet\Desktop\Testing\inputfile.csv"
 
# Specify the path to the shared folder
$sharedFolderPath = "C:\Users\Cablet\Desktop\Testing"
 
# Specify the path to the output file
$outputFilePath = "C:\Users\Cablet\Desktop\VENA TOOL\File.csv"
 
$csvData = Import-Csv -Path $inputCsvPath

# Initialize an array to hold the output data
$outputData = @()
 
foreach ($row in $csvData) {
    $userInput = $row.Name
    $inputFound = $false
 
    Get-ChildItem -Path $sharedFolderPath -Recurse | ForEach-Object {
        if ($_.Extension -eq ".zip") {
            $tempFolder = [System.IO.Path]::GetTempFileName()
            Remove-Item -Path $tempFolder
            New-Item -ItemType Directory -Path $tempFolder | Out-Null
            Expand-Archive -Path $_.FullName -DestinationPath $tempFolder
 
            Get-ChildItem -Path $tempFolder -Recurse -File | ForEach-Object {
                if ($_.Extension -eq ".xlsx") {
                    $excelData = Import-Excel -Path $_.FullName
 
                    $matches = $excelData | Where-Object { $_ -match $userInput }
 
                    if ($matches) {
                        foreach ($match in $matches) {
                            if ($match -match "Pre-Authenticati") {
                                $output = "Found '$userInput': $match"
                                Write-Host $output
                                $outputData += New-Object PSObject -Property @{
                                    'UserName' = $userInput
                                    'Workstation' = ($match -split ";")[2]
                                    'Time' = ($match -split ";")[0]
                                    'AccountLockout' = ($match -split ";")[3]
                                }
                                $inputFound = $true
                            }
                        }
                    }
                } else {
                    $matches = Select-String -Path $_.FullName -Pattern $userInput
                    if ($matches) {
                        foreach ($match in $matches) {
                            if ($match.Line -match "Pre-Authenticati") {
                                $output = "Found '$userInput': $($match.Line)"
                                Write-Host $output
                                $outputData += New-Object PSObject -Property @{
                                    'UserName' = $userInput
                                    'Workstation' = ($match.Line -split ";")[2]
                                    'Time' = ($match.Line -split ";")[0]
                                    'AccountLockout' = ($match.Line -split ";")[3]
                                }
                                $inputFound = $true
                            }
                        }
                    }
                }
            }
 
            Remove-Item -Path $tempFolder -Recurse
        } else {
            $matches = Select-String -Path $_.FullName -Pattern $userInput
            if ($matches) {
                foreach ($match in $matches) {
                    if ($match.Line -match "Pre-Authenticati") {
                        $output = "Found '$userInput': $($match.Line)"
                        Write-Host $output
                        $outputData += New-Object PSObject -Property @{
                            'UserName' = $userInput
                            'Workstation' = ($match.Line -split ";")[2]
                            'Time' = ($match.Line -split ";")[0]
                            'AccountLockout' = ($match.Line -split ";")[3]
                        }
                        $inputFound = $true
                    }
                }
            }
        }
    }
 
    if (-not $inputFound) {
        $output = "The input '$userInput' was not found in any file in the shared folder."
         Write-Host $output
        
     }
}

# Export the output data to a CSV file
$outputData | Export-Csv -Path $outputFilePath -NoTypeInformation
