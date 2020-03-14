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

Get-ChildItem "C:\Users\Noam\Music\Pop" | 
Foreach-Object {
if ((Test-FileLock $_.FullName) -eq $True) {
Write-Host ($_.FullName)
}
}