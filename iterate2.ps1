function Test-FileLock {
  param (
    [parameter(Mandatory=$true)][string]$Path
  )

  $oFile = New-Object System.IO.FileInfo $Path

  if ((Test-Path -Path $Path) -eq $false) {
    return $false
  }

  try {
    $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)

    if ($oStream) {
      $oStream.Close()
    }
    return $false
  } catch {
    # file is locked by a process.
    return $true
  }
}

function presentationText($targetslide, $text) {
	$s = $presentation.Slides[$targetslide]
	$f0a=$s.Shapes[5].TextFrame.TextRange.Text
	$f1a = $text
	$test = $s.Shapes[5].TextFrame.TextRange.Replace($f0a, $f1a)
}

Get-ChildItem "C:\Users\Noam\Music\Pop" | 
Foreach-Object {
if ((Test-FileLock $_.FullName) -eq $True) {
#Write-Host ($_.FullName)
presentationText("../printFile.pptx", $_.FullName)
}
}