Function Export-Excel
{
Param
(
    [string] $excelFilePath,
    [DateTime]$day,
    [string]$drug, 
    [string]$dose,
    [int]$startHour,
    [int]$startMinute, 
    [int]$timeToSleep,
    [int]$wakeHour,
    [int]$wakeMinute,
    [int]$endHour,
    [int]$endMinute,
    [int]$restfulMins,
    [int]$restlessMins,
    [int]$longestMins,
    [int]$sleepQuotient
)

# http://sqlmag.com/powershell/update-excel-spreadsheets-powershell

    Set-Variable -Name xlCellTypeLastCell -Value 11 -Option Constant

    $Excel = New-Object -ComObject Excel.Application
    $ExcelWorkbook = $Excel.Workbooks.Open($excelFilePath)

    # There must be a tab named "SleepIq"
    $ExcelWorksheet = $Excel.WorkSheets.item("SleepIq")
    $ExcelWorksheet.activate()

    # This method of getting the next available row will
    # include empty looking rows the user modified
    # To remove them, delete the rows, don't use clear
    $used = $ExcelWorksheet.usedRange

    $lastCell = $used.SpecialCells($xlCellTypeLastCell)
    $nextRow = $lastCell.Row + 1


    # A, Weekday (from Column B)
    $ExcelWorksheet.Cells.Item($nextRow,1) = "=WEEKDAY(B$($nextRow))"
    # B, Sleep Date
    $ExcelWorksheet.Cells.Item($nextRow,2) = $day.Date.ToString("d")
    # C NapFrom
    # D NapTo
    # E Drug
    $ExcelWorksheet.Cells.Item($nextRow,5) = $drug
    # F Dose
    $ExcelWorksheet.Cells.Item($nextRow,6) = $dose
    # G Bedtime
    $ExcelWorksheet.Cells.Item($nextRow,7) = "=TIME($startHour,$startMinute,0)"
    # H Time to sleep
    $ExcelWorksheet.Cells.Item($nextRow,8) = $timeToSleep
    # I Wake Time
    $ExcelWorksheet.Cells.Item($nextRow,9) = "=TIME($wakeHour, $wakeMinute,0)"
    # J Rise Time
    $ExcelWorksheet.Cells.Item($nextRow,10) = "=TIME($endHour,$endMinute,0)"
    # K Restful time
    $ExcelWorksheet.Cells.Item($nextRow,11) = "=TIME( $([int][Math]::Truncate($restfulMins/60)), $([int]$restfulMins%60), 0)"
    # L Restless time
    $ExcelWorksheet.Cells.Item($nextRow,12) = "=TIME( $([int][Math]::Truncate($restlessMins/60)), $([int]$restlessMins%60), 0)"
    # M Longest sleep
    $ExcelWorksheet.Cells.Item($nextRow,13) = "=TIME( $([int][Math]::Truncate($longestMins/60)), $([int]$longestMins%60), 0)"
    # N Sleep Score
    $ExcelWorksheet.Cells.Item($nextRow,14) = $sleepQuotient
    # O Sleep Efficient Formula Restful/Restful+Restless * 100
    $ExcelWorksheet.Cells.Item($nextRow,15) = "=(K$($nextRow)*60*24)/((K$($nextRow)*60*24)+(L$($nextRow)*60*24))*100"


    $ExcelWorkbook.Save()
    $ExcelWorkbook.Close($true)
}

Export-ModuleMember -Function Export-Excel
