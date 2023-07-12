<#
STEPS
>1. Get downloaded CSV file from User (filter them by shipping type - Order Placed)
>2. Import file to powershell
>3. Check if file is valid (if not try again)
4. Get specific columns from the CSV
5. Place them in new sheet
6. Add semi-colon to every elements end
7. Merge them all into one cell in this format: To;Message;From;Id
8. Export the CSV file in a destination
#>



#-------------------------------------------------------------------------------------------#
#Open CSV File#

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter           = 'CSV Files (*.csv)|*.csv'
    Title            = 'Select a CSV file for Giftcard design'
}
#Prompt a folder selection dialog box to the user
$null = $FileBrowser.ShowDialog()
$FilePath = $FileBrowser.FileName

# Check if file path valid
if (Test-Path $FilePath) {



    #CSV file imported
    #-------------------------------------------------------------------------------------------#
    #Filter the CSV file to keep the rows that have a specific value
    
    $data = Import-Csv $FilePath | where-object Status -eq 'Order Placed'

    #-------------------------------------------------------------------------------------------#
    #Format the Properties in a way that it makes sense to our Giftcard design autoamtion jsx

    $data = $data | Select-Object 'Message To', 'Gift Message', 'Message From', 'Id' | ForEach-Object {
        foreach ($property in $_.PSObject.Properties) {
            $property.Value = $property.Value.ToString() + ';'

            $_.Id = $_.Id -replace '\.0$', ''
            $_.Id = $_.Id.Substring(0, 4) + ';'
        }
        $_.'Gift Message' = $_.'Gift Message' -replace "`r`n", " " -replace "`n", "" # replace newline with space
        $_.'Message To' = $_.'Message To' + $_.'Gift Message' + $_.'Message From' + $_.'Id'
        $_.'Gift Message' = ""
        $_.'Message From' = ""
        $_.'Id' = ""

        $_
    }
    
    $data = $data | Select-Object -Property @{ Name = ';;;;'; Expression = { $_.'Message To' } }

    
    # Create the output directory in the same folder as the PowerShell script
    $outputDirectory = Join-Path -Path $PSScriptRoot -ChildPath ("Output")
    New-Item -Path $outputDirectory -ItemType Directory -Force

    # Export
    $outputPath = Join-Path -Path $outputDirectory -ChildPath "formatted-csv.csv"
($data | ConvertTo-Csv -NoTypeInformation)[1..($data.count)] | Set-Content -Path $outputPath
    Write-Host "CSV file updated and saved to: $outputPath"



}

else {
    Write-Host "Invalid file path. Please select a valid CSV file."
}


#CSV file exported
#-------------------------------------------------------------------------------------------#
#Launch Photoshop and execute script

& cscript.exe 'runPhotoshopScript.vbs'

