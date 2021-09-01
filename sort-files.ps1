Clear-Host

# d druk
$doDruku = "DoDruku"
$doDrukuPath = Test-Path -Path $doDruku
# o obrobka
$doObrobki = "DoObrobki"
$doObrobkiPath = Test-Path -Path $doObrobki
# s strona
$doStrony = "DoStrony"
$doStronyPath = Test-Path -Path $doStrony
# i inne
$doInne = "DoInne"
$doInnePath = Test-Path -Path $doInne

if ($args.Length -lt 1) {
  Write-Output "Need entry file."
  exit 1
}

try {

  if (-Not $doDrukuPath) {
    New-Item . -Name $doDruku -ItemType "directory"
  } 
  if (-Not $doObrobkiPath) {
    New-Item . -Name $doObrobki -ItemType "directory"
  } 
  if (-Not $doStronyPath) {
    New-Item . -Name $doStrony -ItemType "directory"
  }
  if (-Not $doInnePath) {
    New-Item . -Name $doInne -ItemType "directory"
  } 

  $arrayFromFile = Get-Content -Path @($args[0])
  $arrayFromFile = $arrayFromFile[1..($arrayFromFile.Length - 1)]

  foreach ($item in $arrayFromFile) {
    $oneItem = $item.split(";")
    if ($oneItem[1] -eq "d") {
      Move-Item -Path $oneItem[0] -Destination $doDruku
    }
    elseif ($oneItem[1] -eq "o") {
      Move-Item -Path $oneItem[0] -Destination $doObrobki
    }
    elseif ($oneItem[1] -eq "s") {
      Move-Item -Path $oneItem[0] -Destination $doStrony
    }
    elseif ($oneItem[1] -eq "i") {
      Move-Item -Path $oneItem[0] -Destination $doInne
    }
  }
  exit 0
}
catch { 
  "Task failed successfully." 
  exit 1
}