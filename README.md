# Filter Tickets script

## Installing
1. Please ensure that Powershell is enabled
2. Download the script and the 'setup.bat' file.
3. Run 'setup.bat' and provide admin credentials when prompted. This only needs to be run once.

## Configuring
1. Edit the 'filterTickets.ps1' file. Look for the line that begins `$agents = @{`, and modify the list. Each agent name should be written exactly as it appears in ServiceNow, and each line must end in a comma except the last one.
2. In case you wish to modify the output format, please look for the line that begins `$out = "$agent - $date.csv"` and modify it to fit your needs. Examples below:
  * `$out = "$date - $agent.csv"` - creates a folder for each agent, saves the CSV in that folder with the date as the name.
  * `$out = "Tickets handled by $agent on $date.csv"` - slightly more verbose version of the default.
3. Set up a report in ServiceNow that can be exported to CSV. You can choose what is in this report, and choose which columns are displayed, as long as the column 'Work notes' is present.

## Running
1. Run the report in ServiceNow and export it to CSV
2. Copy the report into the script's root folder
3. Note the name of the report (eg: incident.csv)
4. Right click on 'filterTickets.ps1', and choose 'Run with PowerShell'
5. Provide the date you wish to filter for in the format used by ServiceNow in the report (eg: 16/11/2015)
6. Provide the name of the input file (eg: incident.csv)
7. The resulting CSV files will be created for each agent in the folder containing the script

## FAQ
1. Why does 'setup.bat' need admin credentials?
  * Powershell by default ships with restrictions that prevent running unsigned scripts. As filterTickets.ps1 is not signed, the ExecutionPolicy needs to be modified on the computer to allow unsigned scripts to run. There is a minor security risk associated with this, but provided you do not double-click on things that you don't know what they will do, the risk is minimal.

2. Can I export to folders?
  * Not in the current version, as all forward slashes in the path will be replaced with a space to avoid Powershell errors.


3. Can I run the script directly from PowerShell?
  * Sure, just run `.\filterTickets.ps1`

4. Does the script accept any arguments?
  * Yes, the script can be run with the `-in` argument, which takes the name of the input file, and the `-date` argument, which takes the date. Please be sure to wrap your arguments in quotation marks (eg: `.\filterTickets.ps1 -in "incident.csv" -date "16/11/2015"`)