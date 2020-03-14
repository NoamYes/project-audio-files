
param([string]$filePath="")


# $presentation.Slides.add(2), ppLayoutBlank
# Write-Host ($presentation.Slides)
$mydocument = $presentation.Slides.Item(1)

	$f0a=$mydocument.Shapes[1].TextFrame.TextRange.Text
	$f1a = $filePath
	$test = $mydocument.Shapes[1].TextFrame.TextRange.Replace($f0a, $f1a)
# $mydocument.Shapes.title.TextFrame.TextRange.Text = "test"