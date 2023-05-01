$sourceFilePathVert = 'templates\Vert Erbessd Parcing Tool.xlsm'
$destinationFilePathVert = "output\Vert\Vert.xlsm" 

Copy-Item $sourceFilePathVert $destinationFilePathVert -Force
Set-ItemProperty -Path $destinationFilePathVert -Name IsReadOnly -Value $false

$sourceFilePathHoriz = 'templates\Horiz Erbessd Parcing Tool.xlsm'
$destinationFilePathHoriz = "output\Horiz\Horiz.xlsm"


Copy-Item $sourceFilePathHoriz $destinationFilePathHoriz -Force
Set-ItemProperty -Path $destinationFilePathHoriz -Name IsReadOnly -Value $false

$sourceFilePathAxial = 'templates\Axial Erbessd Parcing Tool.xlsm'
$destinationFilePathAxial = "output\Axial\Axial.xlsm"

Copy-Item $sourceFilePathAxial $destinationFilePathAxial -Force
Set-ItemProperty -Path $destinationFilePathAxial -Name IsReadOnly -Value $false

$sourceFilePath = 'templates\Alarm Calculator Template.xlsm'
$destinationFilePath = "output\alarm calculator\ alarm calculator.xlsm"
Copy-Item $sourceFilePath $destinationFilePath -Force
Set-ItemProperty -Path $destinationFilePath -Name IsReadOnly -Value $false



$folder = 'Values'
$folder2= 'output\alarm calculator'
$fileList = Get-ChildItem -Path $folder
$fileList2 = Get-ChildItem -Path $folder2
$excelFiles = $fileList | Where-Object {$_.Extension -eq '.xlsx' -or $_.Extension -eq '.xls'}
$excelFiles2 = $fileList2 | Where-Object {$_.Extension -eq '.xlsm' -or $_.Extension -eq '.xls'}
$excel = New-Object -ComObject Excel.Application
foreach ($file in $excelFiles) {
    $workbook = $excel.Workbooks.Open($file.FullName)
    $destFilePath2 = $excelFiles2 | Sort-Object LastWriteTime | Select-Object -Last 1
    $destworkbook2 = $excel.Workbooks.Open($destFilePath2.FullName)
    $worksheet = $workbook.Worksheets.Item("Sheet1")
    $value = $worksheet.Cells.Item(2, 6).Value2
    $cellValue = $worksheet.Range("A1:J400").Value2
    
    if ($value -eq 1) {
        Write-Output "h"
        $folder = 'output\Horiz\'  
        $fileList = Get-ChildItem -Path $folder
        $excelFiles = $fileList | Where-Object {$_.Extension -eq '.xlsm' -or $_.Extension -eq '.xls'}

        # Get the destination workbook
        $destFilePath = $excelFiles | Sort-Object LastWriteTime | Select-Object -Last 1
        Write-Host "The path of the oldest Excel file is $($destFilePath.FullName)"

        $destworkbook1 = $excel.Workbooks.Open($destFilePath.FullName)
        
        $destworksheet = $destWorkbook1.Worksheets.Item("Raw_Data")

        # Paste the cell value into the destination workbook
        $destworksheet.Range("A1:J400").Value2 = $cellValue
        $excel.Run("Sheet2.Parce_Data_PB_Click")
        $destWorkbook1.save()
        

        $worksheet9010 = $destWorkbook1.Worksheets.Item("9010-ACCEL-CF-g")
        $worksheet9038 = $destWorkbook1.Worksheets.Item("9038-ACC-D-P-Derived-PEAK-ACCEL")
        $worksheet9046 = $destWorkbook1.Worksheets.Item("9046-ACCEL-P-P-Peak-to-Peak-ACC")
        $worksheet0 = $destWorkbook1.Worksheets.Item("0-ACCEL-RMS-g")
        $worksheet9042 = $destWorkbook1.Worksheets.Item("9042-ACCEL-T-P-True-Peak-ACC")


        $cellValue21 =$worksheet9010.Range("A1:B22").Value2
        $cellValue22 =$worksheet9038.Range("A1:B22").Value2
        $cellValue23 =$worksheet9046.Range("A1:B22").Value2
        $cellValue24 =$worksheet0.Range("A1:B22").Value2
        $cellValue25 =$worksheet9042.Range("A1:B22").Value2

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend1")
        $destworksheet2.range("A1:B22").Value2 = $cellValue21

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend4")
        $destworksheet2.range("A1:B22").Value2 = $cellValue22

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend7")
        $destworksheet2.range("A1:B22").Value2 = $cellValue23

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend10")
        $destworksheet2.range("A1:B22").Value2 = $cellValue24

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend14")
        $destworksheet2.range("A1:B22").Value2 = $cellValue25
        
        $destWorkbook2.save()

        $destworkbook2.close()


        $destWorkbook1.Close()

        # Quit the Excel application
        $excel.Quit()
    }
    elseif ($value -eq 2) {
        Write-Output "v"
        $folder = 'output\Vert\'
        $fileList = Get-ChildItem -Path $folder
        $excelFiles = $fileList | Where-Object {$_.Extension -eq '.xlsm' -or $_.Extension -eq '.xls'}

        # Get the destination workbook
        $destFilePath = $excelFiles | Sort-Object LastWriteTime | Select-Object -Last 1
        Write-Host "The path of the oldest Excel file is $($destFilePath.FullName)"

        $destWorkbook1 = $excel.Workbooks.Open($destFilePath.FullName)
        
        $destworksheet1 =  $destWorkbook1.Worksheets.Item("Raw_Data")

        # Paste the cell value into the destination workbook
        $destworksheet1.Range("A1:J400").Value2 = $cellValue
        $excel.Run("Sheet2.Parce_Data_PB_Click")
        $destWorkbook1.save()
        

        $worksheet9010 = $destWorkbook1.Worksheets.Item("9010-ACCEL-CF-g")
        $worksheet9038 = $destWorkbook1.Worksheets.Item("9038-ACC-D-P-Derived-PEAK-ACCEL")
        $worksheet9046 = $destWorkbook1.Worksheets.Item("9046-ACCEL-P-P-Peak-to-Peak-ACC")
        $worksheet0 = $destWorkbook1.Worksheets.Item("0-ACCEL-RMS-g")
        $worksheet9042 = $destWorkbook1.Worksheets.Item("9042-ACCEL-T-P-True-Peak-ACC")


        $cellValue21 =$worksheet9010.Range("A1:B22").Value2
        $cellValue22 =$worksheet9038.Range("A1:B22").Value2
        $cellValue23 =$worksheet9046.Range("A1:B22").Value2
        $cellValue24 =$worksheet0.Range("A1:B22").Value2
        $cellValue25 =$worksheet9042.Range("A1:B22").Value2

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend2")
        $destworksheet2.range("A1:B22").Value2 = $cellValue21

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend5")
        $destworksheet2.range("A1:B22").Value2 = $cellValue22

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend8")
        $destworksheet2.range("A1:B22").Value2 = $cellValue23

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend11")
        $destworksheet2.range("A1:B22").Value2 = $cellValue24

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend15")
        $destworksheet2.range("A1:B22").Value2 = $cellValue25
        
        $destWorkbook2.save()
        


        $destWorkbook1.Close()

        # Quit the Excel application
        $excel.Quit()
    }
    elseif ($value -eq 3) {
        Write-Output "a"
        $folder = 'output\Axial\'
        $fileList = Get-ChildItem -Path $folder
        $excelFiles = $fileList | Where-Object {$_.Extension -eq '.xlsm' -or $_.Extension -eq '.xls'}

        # Get the destination workbook
        $destFilePath = $excelFiles | Sort-Object LastWriteTime | Select-Object -Last 1
        Write-Host "The path of the oldest Excel file is $($destFilePath.FullName)"

        $destWorkbook1 = $excel.Workbooks.Open($destFilePath.FullName)
        
        $destworksheet1 =  $destWorkbook1.Worksheets.Item("Raw_Data")

        # Paste the cell value into the destination workbook
        $destworksheet1.Range("A1:J400").Value2 = $cellValue
        $excel.Run("Sheet2.Parce_Data_PB_Click")
        $destWorkbook1.save()
        

        $worksheet9010 = $destWorkbook1.Worksheets.Item("9010-ACCEL-CF-g")
        $worksheet9038 = $destWorkbook1.Worksheets.Item("9038-ACC-D-P-Derived-PEAK-ACCEL")
        $worksheet9046 = $destWorkbook1.Worksheets.Item("9046-ACCEL-P-P-Peak-to-Peak-ACC")
        $worksheet0 = $destWorkbook1.Worksheets.Item("0-ACCEL-RMS-g")
        $worksheet9042 = $destWorkbook1.Worksheets.Item("9042-ACCEL-T-P-True-Peak-ACC")


        $cellValue21 =$worksheet9010.Range("A1:B22").Value2
        $cellValue22 =$worksheet9038.Range("A1:B22").Value2
        $cellValue23 =$worksheet9046.Range("A1:B22").Value2
        $cellValue24 =$worksheet0.Range("A1:B22").Value2
        $cellValue25 =$worksheet9042.Range("A1:B22").Value2

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend3")
        $destworksheet2.range("A1:B22").Value2 = $cellValue21

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend6")
        $destworksheet2.range("A1:B22").Value2 = $cellValue22

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend9")
        $destworksheet2.range("A1:B22").Value2 = $cellValue23

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend12")
        $destworksheet2.range("A1:B22").Value2 = $cellValue24

        $destworksheet2 = $destworkbook2.Worksheets.Item("Trend16")
        $destworksheet2.range("A1:B22").Value2 = $cellValue25
        
        $destWorkbook2.save()
        

        $destWorkbook1.Close()
    }
    $workbook.Close()
}
