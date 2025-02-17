# Connect to Exchange Online PowerShell session
Connect-ExchangeOnline

# Set the path to your CSV file
$csvFile = "C:\Users\PLACEHOLDER\Downloads\DistributionGroups (3).csv"
# Set the path to the output CSV file
$outputCsvFile = "C:\Users\PLACEHOLDER\Downloads\output.csv"

# Read the CSV file
$data = Import-Csv -Path $csvFile

# Create an array to store DL members for CSV export
$csvData = @()

# Iterate through each DL in the CSV
foreach ($row in $data) {
    $dlName = $row.name

    # Get the DL members
    $members = Get-DistributionGroupMember -Identity $dlName | Select-Object -ExpandProperty PrimarySmtpAddress

    # Display the DL name and its members
    Write-Output "DL: $dlName"
    Write-Output "Members:"
    $members | ForEach-Object {
        Write-Output "- $_"
    }
    Write-Output ""

    # Add the DL members to the CSV data
    $csvData += [PSCustomObject]@{
        DL = $dlName
        Members = $members -join ", "
    }
}

# Export the CSV data to a CSV file
$csvData | Export-Csv -Path $outputCsvFile -NoTypeInformation

