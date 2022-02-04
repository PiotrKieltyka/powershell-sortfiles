Clear-Host

$path = Get-Location
$filename = 'PhotosList.xlsx'
$saveAs = -join ($path, '\', $filename)
Write-Host $saveAs

if (Test-Path $filename) { Remove-Item $filename -Force -ErrorAction SilentlyContinue }

try
{	
	$excel = New-Object -ComObject Excel.Application.16
	$excel.visible = $false
	$workbook = $excel.Workbooks.Add()
	$worksheet = $workbook.Worksheets.Item(1)
	
	$worksheet.Name = 'Fotografia Betlinscy'
	$worksheet.Cells.Item(1, 1).ColumnWidth = 20
	$worksheet.Cells.Item(1, 2).ColumnWidth = 15
	$worksheet.Cells.Item(1, 3).ColumnWidth = 15
	$worksheet.Cells.Item(1, 4).ColumnWidth = 15
	$worksheet.Cells.Item(1, 5).ColumnWidth = 15
	$worksheet.Cells.Item(1, 1) = 'Nazwa pliku'
	$worksheet.Cells.Item(1, 2) = 'd - druk'
	$worksheet.Cells.Item(1, 3) = 'o - obrobka'
	$worksheet.Cells.Item(1, 4) = 's - strona'
	$worksheet.Cells.Item(1, 5) = 'i - inne'
	
	$photosListing = Get-ChildItem -Path . -Name -Include *.jpg, *.jpeg, *.raw
	$cellNumber = 2
	foreach ($photo in $photosListing)
	{
		$worksheet.Cells.Item($cellNumber, 1) = $photo
		$cellNumber += 1
	}
	$worksheet.Columns('B').Locked = $false
	$worksheet.Columns('C').Locked = $false
	$worksheet.Columns('D').Locked = $false
	$worksheet.Columns('E').Locked = $false
	$worksheet.Protect('darekbet')
	
	#Write-Output $worksheet
	
	$workbook.SaveAs($saveAs)
	$excel.Quit()
	Write-Output 'Plik zostal wygenerowany:' $filename
	exit 0
}

catch
{
	Write-Error 'Zadanie posypalo sie z sukcesem.'
	exit 1
}