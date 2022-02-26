Clear-Host

$path = Get-Location
$entryFile = 'PhotosList.xlsx'
$openFile = -join ($path,'\',$entryFile)
Write-Host $openFile

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

  $excel = New-Object -ComObject Excel.Application.16
	$excel.visible = $false
  $workbook = $excel.Workbooks.Open($openFile,$null,$true)
  $worksheet = $workbook.Worksheets.Item(1)
  $lastrow = $worksheet.UsedRange.rows.count + 1
  $counter = 0

  for ($i = 2; $i -lt $lastrow; $i++) {
    if (-Not (Test-Path $worksheet.Cells.Item($i, 1).Value2)) {
      Write-Host 'W tym katalogu brakuje zdjecia:' $worksheet.Cells.Item($i, 1).Value2
      Write-Host 'A zdjecie znajduje sie na liscie:' $entryFile
    } else {
      if ($worksheet.Cells.Item($i, 2).Value2 -eq 'd') {
        Copy-Item -Path $worksheet.Cells.Item($i, 1).Value2 -Destination $doDruku
        $counter++
      }
      if ($worksheet.Cells.Item($i, 3).Value2 -eq 'o') {
        Copy-Item -Path $worksheet.Cells.Item($i, 1).Value2 -Destination $doObrobki
        $counter++
      }
      if ($worksheet.Cells.Item($i, 4).Value2 -eq 's') {
        Copy-Item -Path $worksheet.Cells.Item($i, 1).Value2 -Destination $doStrony
        $counter++
      }
      if ($worksheet.Cells.Item($i, 5).Value2 -eq 'i') {
        Copy-Item -Path $worksheet.Cells.Item($i, 1).Value2 -Destination $doInne
        $counter++
      }
    }
  }

  $excel.Quit()
  
  Write-Host 'Zadanie jak zwykle zakonczone sukcesem.'
  Write-Host 'Skopiowanych plikow w sumie:' $counter
  exit 0
}
catch { 
  'Zadanie jak zwykle posypalo sie z sukcesem.'
  exit 1
}