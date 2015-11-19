## FilterTickets - Joe Innes 2015 (joseph.innes@tcs.com)
## 
## Script to run against ServiceNow exports to identify all tickets an agent has touched during a single day based on updates to the work notes
##
## Usage: .\filterTickets.ps1 -date <date in DD/MM/YYYY format> -in <inputfile.csv>
##
## You may have any number of columns, exported as a CSV from ServiceNow
## If date and input file are not provided as arguments, they will be prompted for.
##
## !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
## !! Do not touch the next four lines !!
## !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

param (
    [string]$date = $( Read-Host "Date" ), # Prompts the user for the date if this is not provided as an argument
    [string]$in = $( Read-Host "Input file" ) # Prompts the user for the input file if this is not provided as an argument
)


## The output filename is by default "<Agent name> - DD/MM/YYYY.csv". This can be customised by choosing the first
## Customise the lines below to add agents. Make sure each line EXCEPT the last ends with a comma.

$agents = @(
    "Gabor Csecsur",
    "Yasser Puentes",
    "Lilian Halasz"
) 

$csv = Import-Csv $in # Import the input file from a CSV into a Powershell Object

foreach ($agent in $agents) { # For each agent in the list

## Modify the line below to change the output file naming convention

    $out = "$agent - $date.tmp" # Set up the agent's output file name.


    $out = $out -replace "\/", " " # Replace the / in the date because Windows won't let you use that in a file name
    $filteredCsv = $csv | Where-Object {$_.work_notes -like "*$date ??:??:?? - $agent*"} # The heavy lifting - fetch each row where the Worknotes column contains "DD/MM/YYYY HH:MM:SS - Agent Name" and store it in a new var
    $filteredCsvPlusName = $filteredCsv | Select-Object *,@{Name='Agent';Expression={$agent}}
    $filteredCsvPlusName | Export-Csv $out -notype -Encoding "UTF8"

}


$getFirstLine = $true # Get the first line

## For each .tmp file
get-childItem "*.tmp" | foreach {

    if ($filePath -ne $in)
    {

      $filePath = $_

      $lines =  $lines = Get-Content $filePath  # Get the the lines out of the file
      $linesToWrite = switch($getFirstLine) { 
           $true  {$lines} # Put all of them in the $linesToWrite var if getfirstline is true
           $false {$lines | Select -Skip 1} # Otherwise, skip the first line

      }

      $getFirstLine = $false # Don't get the first line again
      Add-Content "merged.csv" $linesToWrite # Output the final CSV
      Remove-Item $_ # Delete the temp file
    }
}