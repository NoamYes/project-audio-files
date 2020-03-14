Add-Type -AssemblyName office

$Application = New-Object -ComObject PowerPoint.Application

$Application.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

$slideType = "microsoft.office.interop.powerpoint.ppSlideLayout" -as [type]

$templatePresentation = "C:\Users\Noam\OneDrive - Technion\Other Courses\Projects\printFile.pptx"

$presentation = $Application.Presentations.open($templatePresentation)

# $presentation.Slides.add(2), ppLayoutBlank
# Write-Host ($presentation.Slides)
$mydocument = $presentation.Slides.Item(1)

	$f0a=$mydocument.Shapes[1].TextFrame.TextRange.Text
	$f1a = "Hello world"
	$test = $mydocument.Shapes[1].TextFrame.TextRange.Replace($f0a, $f1a)
# $mydocument.Shapes.title.TextFrame.TextRange.Text = "test"