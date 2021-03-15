# & cls & powershell -Command "Invoke-Command -ScriptBlock ([ScriptBlock]::Create(((Get-Content """%0""") -join """`n""")))" & exit
# The above line makes the script executable when renamed .cmd or .bat

# Show and export all printers
# Author: Anton Pusch (Stud-IT) 
# Last update: 2020-11-23

# All data is saved to a user folder in the root directory
$folder = $PWD.Drive.Root + $env:username

# Create folder if necessary and open it
if (Test-Path -Path $folder) {
    Write-Host "Using directory $folder`n"
} else {
    Write-Host "Create directory $folder`n"
    New-Item -ItemType Directory -Path $folder | Out-Null
}
Set-location -Path $folder

# List all installed printers and export them to a spreadsheet

# Get list of all printers
Write-Host "Scanning through printers... " -NoNewLine
$printers = Get-Printer
Write-Host ("Found " + $printers.Count + " printers")

# Display all printers grouped in type
foreach ($group in $printers | group Type) {
   Write-Output ($printers | Where-Object {$_.Type -eq $group.Name})
   Write-Host "`n"
}


# | select -Property Type, Name, PortName, Comment, Location, DriverName

# Get list of all drivers
Write-Host "`nScanning through drivers... " -NoNewLine
$drivers = @{}
foreach ($item in $printers) {
   # create hashtable of all used drivers
   $drivers[$item.DriverName] = Get-PrinterDriver -Name $item.DriverName
}
Write-Host ("Found " + $drivers.Count + " drivers") -NoNewLine
Write-Output ($drivers | Format-Table)

# Merge printer and driver properties into printers array
$printerProperties = Get-Printer | Get-Member -MemberType Properties
$driverProperties = Get-PrinterDriver | Get-Member -MemberType Properties

foreach ($property in $driverProperties)
{
   if ($printerProperties.Name.Contains($property.Name)) 
   {
      <# Skip property if same name already exists in printers. Effects:
      Caption, CommunicationStatus, ComputerName, Description, DetailedStatus, ElementName, 
      HealthState, InstallDate, InstanceID, Name, OperatingStatus, OperationalStatus, 
      PrimaryStatus, PrintProcessor, PSComputerName, Status, StatusDescriptions #>
      continue
   }

   # Add driver's collums (properties) to printers array
   $printers | Add-Member -MemberType NoteProperty -Name $property.Name -Value $null -Force
   foreach ($item in $printers) {
      # Attach driver details (out of the hashtable) to each printer
      $itemDriver = $drivers[$item.DriverName]
      $item.$($property.Name) = $itemDriver.$($property.Name)
   }
}

# Design output table
$collumns = @() # Output table header
# Prepare a list of collumns and delete the empty ones
$mergedProperties = @("Type", "Name", "PortName", "Comment", "Location", "DriverName", "ConfigFile", "DataFile")
$mergedProperties += ($printers | Get-Member -MemberType Properties).Name
foreach ($property in $mergedProperties) {
   if ($collumns.Contains($property)) {
      # Skip if collumn already exists
      continue
   }
   foreach ($item in $printers) {
      if ($item.$($property) -ne $null) {
         # If at least one printer has a valid value, keep the collumn
         $collumns += $property
         break # Stops collumn duplications
      }
   }
}

# Display all printers and their drivers in a GUI
Write-Host "Opening GUI...`n"
$printers | select $collumns | Out-Gridview -Title "List of all installed printers including their drivers ($folder\printer-$env:username.csv)"

# Export table to Excel
do {
   $error.clear()
   try {
      # Try to override file
      $printers | select $collumns | Export-Csv -Path "printer-$env:username.csv" -Delimiter ";" -NoTypeInformation
   } catch {
      Read-Host -Prompt "Couldn't override file $folder\printer-$env:username.csv`nTry again?`t^C to cancel`tEnter to repeat"
   }
} while ($error)
Write-Host "Exported to $folder\printer-$env:username.csv`n"

# Keep console window open
$reply = Read-Host -Prompt "Finished!`to to open spreadsheet file`tEnter to quit"
if ($reply.ToLower() -eq 'o') {
   # Open the spreadsheet
   Start-Process -FilePath ".\printer-$env:username.csv"
}
