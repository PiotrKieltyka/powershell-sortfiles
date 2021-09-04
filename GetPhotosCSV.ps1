Clear-Host

# Nazwa zapisywanego pliku z lista plikow zdjec w folderze
$fileName = 'PhotosList.csv'

try {
    Remove-Item .\$fileName -ErrorAction SilentlyContinue
    'Nazwa pliku, d druk, o obrobka, s strona, i inne' | Out-File .\$fileName -Append
    $filesListing = Get-ChildItem -Path . -Name -Include *.jpg, *.jpeg, *.raw
    foreach ($file in $filesListing) {
      $file + ',' | Out-File .\$fileName -Append
    }
    Write-Output 'Plik zostal wygenerowany:' $fileName
    exit 0
}
catch {
  'Zadanie posypalo sie z sukcesem.'
  exit 1
}