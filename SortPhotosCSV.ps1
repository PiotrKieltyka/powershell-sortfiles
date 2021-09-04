Clear-Host

# Nazwa odczytywanego pliku z lista plikow zdjec w folderze
$entryFile = 'PhotosList.csv'

if (-Not (Test-Path $entryFile)) {
  'W katalogu brakuje pliku: ' + $entryFile
  exit 1
}

# d druk
$doDruku = 'DoDruku'
# o obrobka
$doObrobki = 'DoObrobki'
# s strona
$doStrony = 'DoStrony'
# i inne
$doInne = 'DoInne'

try {

  if (-Not (Test-Path -Path $doDruku)) {
    [void](New-Item . -Name $doDruku -ItemType 'directory')
  } 
  if (-Not (Test-Path -Path $doObrobki)) {
    [void](New-Item . -Name $doObrobki -ItemType 'directory')
  } 
  if (-Not (Test-Path -Path $doStrony)) {
    [void](New-Item . -Name $doStrony -ItemType 'directory')
  }
  if (-Not (Test-Path -Path $doInne)) {
    [void](New-Item . -Name $doInne -ItemType 'directory')
  } 

  $arrayFromFile = Get-Content -Path $entryFile
  $arrayFromFile = $arrayFromFile[1..($arrayFromFile.Length - 1)]

  $counter = 0

  foreach ($item in $arrayFromFile) {
    $oneItem = $item.split(',')
    if (-Not (Test-Path $oneItem[0])) {
      Write-Host 'W tym katalogu nie ma pliku:' $oneItem[0]
      Write-Host 'A plik jest wpisany w:' $entryFile
    }
    else {
      if ($oneItem[1].Trim().ToLower() -eq 'd') {
        Move-Item -Path $oneItem[0] -Destination $doDruku
        $counter++
      }
      elseif ($oneItem[1].Trim().ToLower() -eq 'o') {
        Move-Item -Path $oneItem[0] -Destination $doObrobki
        $counter++
      }
      elseif ($oneItem[1].Trim().ToLower() -eq 's') {
        Move-Item -Path $oneItem[0] -Destination $doStrony
        $counter++
      }
      elseif ($oneItem[1].Trim().ToLower() -eq 'i') {
        Move-Item -Path $oneItem[0] -Destination $doInne
        $counter++
      }
    }
  }
  Write-Host 'Przeniesionych plikow:' $counter
  exit 0
}
catch { 
  'Zadanie posypalo sie z sukcesem.'
  exit 1
}