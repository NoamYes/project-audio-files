$timeout = new-timespan -Minutes 2
$sw = [diagnostics.stopwatch]::StartNew()
while ($sw.elapsed -lt $timeout){
	clear
	start-sleep -seconds 1
	Measure-Command { & "C:\Users\Noam\OneDrive - Technion\Other Courses\Projects\project-audio-files\iterate2.ps1" }
  
  start-sleep -seconds 3
}

write-host "stop"