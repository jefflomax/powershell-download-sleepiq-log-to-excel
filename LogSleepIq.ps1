Param( [DateTime]$day, [string]$drug, [string]$dose, [string]$user, [string]$pass)

# MM-DD-YYYY
# Usage: .\LogSleepIq.ps1 -Day '07-01-2015' -drug 'OTC' -dose '50' 
# set USER environment variables SleepIqUser & SleepIqPassword
# or pass -user and -pass
# or just hard-code your Sleep IQ Login in the script

$userName = ""
if( $user.Length -ne 0 ){
    $userName = $user
} else {
    $ev = [Environment]::GetEnvironmentVariable("SleepIqUser","User")
    if ( $ev -ne $null ){
        $userName = $ev
    }
}

if ($userName.Length -eq 0 ){
    Write-Host "Must pass -user parameter or set user environment variable SleepIqUser"
    Exit
}

$password = ""
if( $pass.Length -ne 0 ){
    $password = $pass
} else {
    $ev = [Environment]::GetEnvironmentVariable("SleepIqPassword","User")
    if ( $ev -ne $null ){
        $password = $ev
    }    
}

if ($password.Length -eq 0 ){
    Write-Host "Must pass -pass parameter or set user environment variable SleepIqPassword"
    Exit
}

# Import SleepIqModule, Export-Excel
"$pwd\SleepIqModule.psm1", "$pwd\Export-Excel.psm1" | Import-Module

# Powershell destructuring assignment
$iqs, $userId, $key = Start-SleepIqSession -sleepIqUsername $userName -sleepIqPassword $password

$bedInfo = Get-SleepIqBedInfo -iqs $iqs -key $key

Write-Host "Account Id: $($bedInfo.accountId)"
Write-Host "Bed Id: $($bedInfo.bedId)"
Write-Host "Sleeper Left Id: $($bedInfo.sleeperLeftId)"
Write-Host "Sleeper Right Id: $($bedInfo.sleeperRightId)"
Write-Host "Retrieving $($day.ToShortDateString()), SleepIQ stores time from noon to noon in 10 minute intervals"

if( $bedInfo.sleeperLeftId -ne 0 ){

    $spreadsheet = "$pwd\SleeperLeftLog.xlsx"
    
    Write-Host "Found Sleeper Left: $($bedInfo.sleeperLeftId)"
    
    $sleepIqDay = Get-SleepIqSleepDayData -iqs $iqs -key $key -sleeperId $bedInfo.sleeperLeftId -sleepDate $day

    Write-Host "Average Heart Rate: $($sleepIqDay.avgHeartRate)"
    Write-Host "Average Respiration Rate: $($sleepIqDay.avgRespirationRate)"
    Write-Host "Total Sleep Time: $($sleepIqDay.totalSleepSessionTime)"
    Write-Host "Restful Sleep: $($sleepIqDay.restful)"
    Write-Host "Restless Sleep: $($sleepIqDay.restless)"
    
    Write-Host "Restful Minutes: $($sleepIqDay.restfulMins)"
    Write-Host "Restless Minutes: $($sleepIqDay.restlessMins)"
  
    Write-Host "Start DateTime: $($sleepIqDay.startDate)"
    Write-Host "End DateTime: $($sleepIqDay.endDate)"
    Write-Host "Sleep Quotient: $($sleepIqDay.sleepQuotient)"

    $sleepSliceData = Get-SleepIqSleepSliceData -iqs $iqs -key $key -sleeperId $bedInfo.sleeperLeftId -sleepDate $day

    $sleepSliceSummary = Import-SleepIqSleepSliceData -data $sleepSliceData -sleepDate $day

    $lastWake = $sleepIqDay.endDate.AddMinutes( - $sleepSliceSummary.LastRestlessRun )

    Export-Excel -excelFilePath $spreadsheet `
        -day $day `
        -drug $drug `
        -dose $dose `
        -startHour $($sleepIqDay.startDate.ToString("HH")) `
        -startMinute $($sleepIqDay.startDate.ToString("mm")) `
        -timeToSleep $sleepSliceSummary.TimeToSleep `
        -wakeHour $lastWake.ToString("HH") `
        -wakeMinute $lastWake.ToString("mm") `
        -endHour $sleepIqDay.endDate.ToString("HH") `
        -endMinute $sleepIqDay.endDate.ToString("mm") `
        -restfulMins $sleepIqDay.restfulMins `
        -restlessMins $sleepIqDay.restlessMins `
        -longestMins $sleepSliceSummary.LongestRestful `
        -sleepQuotient $sleepIqDay.sleepQuotient

}

if( $bedInfo.sleeperRightId -ne 0 ){

    $spreadsheet = "$pwd\SleeperRightLog.xlsx"

    Write-Host "Found Sleeper Right: $($bedInfo.sleeperRightId)"
   
    $sleepIqDay = Get-SleepIqSleepDayData -iqs $iqs -key $key -sleeperId $bedInfo.sleeperRightId -sleepDate $day

    Write-Host "Average Heart Rate: $($sleepIqDay.avgHeartRate)"
    Write-Host "Average Respiration Rate: $($sleepIqDay.avgRespirationRate)"
    Write-Host "Total Sleep Time: $($sleepIqDay.totalSleepSessionTime)"
    Write-Host "Restful Sleep: $($sleepIqDay.restful)"
    Write-Host "Restless Sleep: $($sleepIqDay.restless)"
    
    Write-Host "Restful Minutes: $($sleepIqDay.restfulMins)"
    Write-Host "Restless Minutes: $($sleepIqDay.restlessMins)"
  
    Write-Host "Start DateTime: $($sleepIqDay.startDate)"
    Write-Host "End DateTime: $($sleepIqDay.endDate)"
    Write-Host "Sleep Quotient: $($sleepIqDay.sleepQuotient)"

    $sleepSliceData = Get-SleepIqSleepSliceData -iqs $iqs -key $key -sleeperId $bedInfo.sleeperRightId -sleepDate $day

    $sleepSliceSummary = Import-SleepIqSleepSliceData -data $sleepSliceData -sleepDate $day

    $lastWake = $sleepIqDay.endDate.AddMinutes( - $sleepSliceSummary.LastRestlessRun )

    Export-Excel -excelFilePath $spreadsheet `
        -day $day `
        -drug $drug `
        -dose $dose `
        -startHour $($sleepIqDay.startDate.ToString("HH")) `
        -startMinute $($sleepIqDay.startDate.ToString("mm")) `
        -timeToSleep $sleepSliceSummary.TimeToSleep `
        -wakeHour $lastWake.ToString("HH") `
        -wakeMinute $lastWake.ToString("mm") `
        -endHour $sleepIqDay.endDate.ToString("HH") `
        -endMinute $sleepIqDay.endDate.ToString("mm") `
        -restfulMins $sleepIqDay.restfulMins `
        -restlessMins $sleepIqDay.restlessMins `
        -longestMins $sleepSliceSummary.LongestRestful `
        -sleepQuotient $sleepIqDay.sleepQuotient
}

Remove-Module SleepIqModule
Remove-Module Export-Excel

