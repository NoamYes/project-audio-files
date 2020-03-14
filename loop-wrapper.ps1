$timeout = new-timespan -Minutes 300
$sw = [diagnostics.stopwatch]::StartNew()

Add-Type -AssemblyName office

$Application = New-Object -ComObject PowerPoint.Application

$Application.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

$slideType = "microsoft.office.interop.powerpoint.ppSlideLayout" -as [type]

$templatePresentation = "C:\Users\Noam\OneDrive - Technion\Other Courses\Projects\printFile.pptx"

$presentation = $Application.Presentations.open($templatePresentation)
while ($sw.elapsed -lt $timeout){
	clear
	start-sleep -seconds 1
	& "C:\Users\Noam\OneDrive - Technion\Other Courses\Projects\project-audio-files\iterate2.ps1" -pptxFile $presentation
  
  start-sleep -seconds 10
}

write-host "stop"