#Author: ianq
#Based off of a monitoring script by: BigTeddy
 
# Since JES outputs local emails as serialized .jmsg files and not .html files
# it seemed high time there was a script that could do it automatically for you.
# This is that script

# To stop the monitoring, run the following commands: 
# Unregister-Event FileCreated 
 
$folder = 'C:\Eclipse\eclipse\plugins\jes.email.server.launcher_1.0.0\jes-1.6.1\smtp' # Enter the root path you want to monitor.
$htmlFolder = 'C:\Eclipse\eclipse\plugins\jes.email.server.launcher_1.0.0\jes-1.6.1\html' # Where the new html files are going
$filter = '*.*' # only .ser files anyhow, change if you want to change other files extensions I guess
 
# In the following line, you can change 'IncludeSubdirectories to $true if required.                           
$fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{IncludeSubdirectories = $false;NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'} 
  
Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action { 
	$name = $Event.SourceEventArgs.Name 
	$changeType = $Event.SourceEventArgs.ChangeType
	$folder = 'C:\Eclipse\eclipse\plugins\jes.email.server.launcher_1.0.0\jes-1.6.1\smtp' # Enter the root path you want to monitor.
	$htmlFolder = 'C:\Eclipse\eclipse\plugins\jes.email.server.launcher_1.0.0\jes-1.6.1\html' # Where the new html files are going
	$timeStamp = $Event.TimeGenerated
	$newName = $name -replace ".ser",".html"

	# Log the event
	Write-Host "New email '$name' was $changeType at $timeStamp" -fore green 
	
	Copy-Item "$folder\$name" $htmlFolder
	Rename-Item -path "$htmlFolder\$name" -newName $newName
}