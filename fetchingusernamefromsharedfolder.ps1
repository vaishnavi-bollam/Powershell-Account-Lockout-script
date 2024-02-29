$userInput = Read-Host -Prompt 'provide Username'

# Specify the path to the shared folder
$sharedFolderPath = "C:\Users\Cablet\Desktop\FCR-csv"

# Specify the path to the output file
$outputFilePath = "C:\Users\Cablet\Desktop\VENA TOOL\File.txt"

$inputFound = $false


Get-ChildItem -Path $sharedFolderPath -Recurse -File | ForEach-Object {
    
    $matches = Select-String -Path $_.FullName -Pattern $userInput
    if ($matches) {
        
        foreach ($match in $matches) {
            $output = "Found '$userInput' in file $($_.FullName) on line $($match.LineNumber): $($match.Line)"
            Write-Host $output
            Add-Content -Path $outputFilePath -Value $output
        }
        $inputFound = $true
    }
}


if (-not $inputFound) {
    $output = "The input '$userInput' was not found in any file in the shared folder."
    Write-Host $output
    Add-Content -Path $outputFilePath -Value $output
}
