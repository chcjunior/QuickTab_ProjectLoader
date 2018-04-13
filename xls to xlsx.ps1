
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault

$excel = New-Object -ComObject excel.application
$excel.visible = $true
$folderpath = "\\SQLTEST11\QuickTab\*"
$filetype ="*xls"
Get-ChildItem -Path $folderpath -Include $filetype | 
ForEach-Object `
{
$path = ($_.fullname).substring(0,($_.FullName).lastindexOf("."))
"Converting $path to $filetype..."
$workbook = $excel.workbooks.open($_.fullname)

$workbook.saveas($path, $xlFixedFormat)
$workbook.close()

$oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\Old xls"
	
	write-host $oldFolder
	if(-not (test-path $oldFolder))
	{
		new-item $oldFolder -type directory
	}
	
	move-item $_.fullname $oldFolder

}
$excel.Quit()
$excel = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()