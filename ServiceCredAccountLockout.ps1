# Specify the path to the shared folder
$sharedFolderPath = "C:\Users\Cablet\Desktop\Testing"

# Specify the path to the output file
$outputFilePath = "C:\Users\Cablet\Desktop\VENA TOOL\File.csv"

# Service account credentials
$serviceAccountUsername = "your_service_account_username"
$serviceAccountPassword = ConvertTo-SecureString "your_service_account_password" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($serviceAccountUsername, $serviceAccountPassword)

# Create a new PSDrive with the service account credentials
New-PSDrive -Name "SharedFolder" -PSProvider FileSystem -Root $sharedFolderPath -Credential $credential

# Now you can access the shared folder through the new PSDrive
$sharedFolderPath = "SharedFolder:\"

# Ask the user for input
$userInput = Read-Host -Prompt 'Enter the lockout user name'
$inputFound = $false

$outputData = @()

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
                                'Workstation' = ($match -split ";")[7]
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
                                'Workstation' = ($match.Line -split ";")[7]
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

$outputData | Export-Csv -Path $outputFilePath -NoTypeInformation
