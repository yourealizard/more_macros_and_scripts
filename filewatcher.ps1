### SET FOLDER TO WATCH + FILES TO WATCH + SUBFOLDERS YES/NO
	$filewatcher = New-Object System.IO.FileSystemWatcher
	#set the folder to monitor
	$filewatcher.Path = "C:\Users\jeNguyen\Desktop\EspressoTechService Company"
	$filewatcher.Filter = "*.*"
	#include subdirectories?
	$filewatcher.IncludeSubdirectories = $false
	$filewatcher.EnableRaisingEvents = $true

### DEFINE ACTIONS AFTER AN EVENT IS DETECTED
	$writeaction = { $path = $Event.SourceEventArgs.FullPath
		$changeType = $Event.SourceEventArgs.ChangeType
		$logline = “$(Get-Date), $changeType, $path”
		Add-content “C:\Users\jeNguyen\Desktop\EspressoTechService Company\FileWatcher_log.txt” -value $logline
		}
### DECIDE WHICH EVENTS SHOULD BE WATCHED
	Register-ObjectEvent $filewatcher “Created” -Action $writeaction
	Register-ObjectEvent $filewatcher “Changed” -Action $writeaction
	Register-ObjectEvent $filewatcher “Deleted” -Action $writeaction
	Register-ObjectEvent $filewatcher “Renamed” -Action $writeaction
	while ($true) {	sleep 5	}
