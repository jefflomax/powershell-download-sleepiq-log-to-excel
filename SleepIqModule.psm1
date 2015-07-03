# For an approved list of verbs:  Get-Verb | Sort-Object Verb

#
# Private Functions
#

# SleepIQ methods take a parameter "_" which is the UNIX timestamp
# milliseconds since 1970.  This probably done to prevent caching.
Function Timestamp
{
    [long](New-Timespan -Start (Get-Date -Date "01/01/1970") -End (Get-Date)).TotalMilliseconds
}

Function UrlDate( [DateTime]$day )
{
    $day.Date.ToString("yyyy-MM-dd")
}

Set-Variable -Name bedInfoHeader `
    -Value @{"Accept"="*/*";"Referer"="https://sleepiq.sleepnumber.com/";"X-Requested-Width"="XMLHttpRequest"} `
    -Option Constant

#
# Public Functions
#

Function Start-SleepIqSession
{
Param ([string]$sleepIqUsername, [string]$sleepIqPassword)

    #Write-Host "Establishing Session for $sleepIqUsername"

    $loginSessionUri = "http://sleepiq.sleepnumber.com/"

    $accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
    $chromeUA = "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.90 Safari/537.36"
    $headers = @{"Accept"=$accept}

    # Access the main website and establish a session in iqs
    $login = Invoke-Webrequest -Uri $loginSessionUri -Headers $headers -UserAgent $chromeUA -SessionVariable iqs

    $body = "{""login"":""$sleepIqUsername"",""password"":""$sleepIqPassword""}"

    # Write-Host $body

    $loginUri = "https://sleepiq.sleepnumber.com/rest/login/?_k="
    $loginInfo = Invoke-Webrequest -Uri $loginUri -METHOD PUT -WebSession $iqs -Body $body

    $jsonLoginInfo = $loginInfo.Content | ConvertFrom-Json

    # userId is [long]
    $userId = $jsonLoginInfo.userId
    # key is [string]
    $key = $jsonLoginInfo.key

    return $iqs, $userId, $key
}


Function Get-SleepIqBedInfo
{
Param( [Microsoft.PowerShell.Commands.WebRequestSession]$iqs, [string]$key )

    [hashtable]$result = @{}

    $bedInfoHeaders = @{"Accept"="*/*";"Referer"="https://sleepiq.sleepnumber.com/";"X-Requested-Width"="XMLHttpRequest"}

    Write-Host "CONST BedInfoHeader: $($bedInfoHeader.Referer)"

    $bedInfoUri = "https://sleepiq.sleepnumber.com/rest/bed/?_=$(Timestamp)&_k=$($key)"
    $bedInfo = Invoke-Webrequest -Uri $bedInfoUri -METHOD GET -WebSession $iqs -Headers $bedInfoHeaders

    $jsonBedInfo = $bedInfo | ConvertFrom-Json

    $result.sleeperLeftId = $jsonBedInfo.beds.sleeperLeftId
    $result.sleeperRightId = $jsonBedInfo.beds.sleeperRightId
    $result.bedId = $jsonBedInfo.beds.bedId
    $result.accountId = $jsonBedInfo.beds.accountId

    return $result
}

Function Get-SleepIqSleepDayData
{
Param( [Microsoft.PowerShell.Commands.WebRequestSession]$iqs, [string]$key, [long]$sleeperId, [DateTime]$sleepDate )

    [hashtable]$result = @{}

    $bedInfoHeaders = @{"Accept"="*/*";"Referer"="https://sleepiq.sleepnumber.com/";"X-Requested-Width"="XMLHttpRequest"}
    $sleepDataUri = "https://sleepiq.sleepnumber.com/rest/sleepData/?date=$(UrlDate($sleepDate))&interval=D1&sleeper=$($sleeperId)&_=$(Timestamp)&_k=$($key)"

    $sleepData = Invoke-Webrequest -Uri $sleepDataUri -METHOD GET -WebSession $iqs -Headers $bedInfoHeaders

    $jsonSleepData = $sleepData | ConvertFrom-Json

    # SleepIq Defined
    $result.avgHeartRate = $jsonSleepData.avgHeartRate 
    $result.avgRespirationRate = $jsonSleepData.avgRespirationRate
    $result.totalSleepSessionTime = $jsonSleepData.totalSleepSessionTime
    $result.restful = $jsonSleepData.restful
    $result.restless = $jsonSleepData.restless
    
    # Summary
    $result.restfulMins = [int]$result.restful / 60
    $result.restlessMins = [int]$result.restless / 60

    $session = $jsonSleepData.sleepData[0].sessions[0]
   
    $result.startDate = Get-Date $session.startDate
    $result.endDate = Get-Date $session.endDate
    $result.sleepQuotient = $session.sleepQuotient

#Write-Output "HR: $avgHeartRate RR: $avgRespirationRate TSST: $(HMST($totalSleepSessionTime)) RFUL: $(HMST($restful)) RLESS: $(HMST($restless)) SD: $startDate ED: $endDate QUO: $sleepQuotient"

    return $result
}

# Retrieve the 10-minute interval sleepDateTime, restfulTime, restType, restlessTime, outOfBedTime
Function Get-SleepIqSleepSliceData
{
Param( [Microsoft.PowerShell.Commands.WebRequestSession]$iqs, [string]$key, [long]$sleeperId, [DateTime]$sleepDate )

    $bedInfoHeaders = @{"Accept"="*/*";"Referer"="https://sleepiq.sleepnumber.com/";"X-Requested-Width"="XMLHttpRequest"}

    $sleepSliceDataUri = "https://sleepiq.sleepnumber.com/rest/sleepSliceData/?date=$(UrlDate($sleepDate))&sleeper=$($sleeperId)&_=$(Timestamp)&_k=$($key)"
    $sleepSlice = Invoke-Webrequest -Uri $sleepSliceDataUri -METHOD GET -WebSession $iqs -Headers $bedInfoHeaders
    $sleepSliceData = $sleepSlice.Content | ConvertFrom-Json


    $array = New-Object System.Collections.ArrayList
    # Slices are noon to noon
    $sliceDateTime = $sleepDate.Date.AddHours(12)

    #Write-Host $sleepSliceData.date
    #Write-Host $sleepSliceData.sleeperId
    #Write-Host $sleepSliceData.sliceSize
    
    foreach( $e in $sleepSliceData.sliceList ){
        $sliceDt = $sliceDateTime.ToString("yyyy-MM-dd HH:mm:ss") # 24 hour time

        # Write-Host """$sliceDt"",$($e.restfulTime),$($e.type),$($e.restlessTime),$($e.outOfBedTime)"
        [hashtable]$slice = @{ 
            dateTime=[DateTime]$sliceDateTime
            restfulTime = $($e.restfulTime)
            type=$($e.type)
            restlessTime=$($e.restlessTime)
            outOfBedTime=$($e.outOfBedTime)
        }
        $array.Add($slice) | Out-Null

        $sliceDateTime = $sliceDateTime.AddMinutes(10)
    }

    return $array
}

Function Import-SleepIqSleepSliceData {
Param( [System.Collections.ArrayList]$data, [DateTime]$sleepDate )

    Set-Variable -Name tNone -Value 0 -Option Constant
    Set-Variable -Name tOutOfBed -Value 1 -Option Constant
    Set-Variable -Name tRestless -Value 2 -Option Constant
    Set-Variable -Name tRestful -value 3 -Option Constant

    #Track first time to sleep
    $firstSleep = $false
    $firstRestfulSleep = $false

    $restfulPeriods = [int[]]@()
    $restfulPeriod = [int]0

    $restlessPeriods = [int[]]@()
    $restlessPeriod = [int]0

    $outOfBedPeriods = [int[]]@()
    $outOfBedPeriod = [int]0
    
    $lastRestlessRun = [int]0

    $firstRestlessTime = $sleepDate.Date.AddHours(12)
    $firstRestfulTime = $sleepDate.Date.AddHours(12)

    $lastType = $tNone

    foreach( $slice in $data ){
        #Write-Host "Time: $($slice.dateTime) Type: $($slice.type) Restful: $($slice.restfulTime) Restless: $($slice.restlessTime) OutOfBed: $($slice.outOfBedTime)"

        if( $slice.type -ne $tNone ){
        
            # Changed type from prior 10 min interval
            if( $slice.type -ne $lastType ){
                
                if( $lastType -eq $tRestful ) {
                    # store as minutes
                    $restfulPeriods += [Math]::Truncate($restfulPeriod/60)
                    $restfulPeriod = 0
                }

                if( $lastType -eq $tRestless ){
                    $restlessPeriods += $restlessPeriod
                    $restlessPeriod = 0
                }

                if( $lastType -eq $tOutOfBed){
                    $outOfBedPeriods += $outOfBedPeriod
                    $outOfBedPeriod = 0
                }

                $lastType = $slice.type
                $lastTypeChange = $slice.dateTime
            }


            #
            # Process each 10 minute interval
            #
            # Accumulate 10 minutes for each restless period
            # (actual time is always far less)
            if( $slice.type -eq $tRestless ){
                $restlessPeriod += 10
                $lastRestlessRun += 10

                if( $firstSleep -eq $false ){
                    $firstSleep = $true
                    $firstRestlessTime = $slice.dateTime
                }
            }

            if( $slice.type -eq $tOutOfBed ){
                $outOfBedPeriod += 10
                $lastRestlessRun += 10
            }

            if( $slice.type -eq $tRestful ){
                $restfulPeriod += $slice.restfulTime
                $lastRestlessRun = 0

                if( $firstRestfulSleep -eq $false ){
                    $firstRestfulSleep = $true
                    $firstRestfulTime = $slice.dateTime
                }
            }
        }# tNone
    }# Foreach Slice

    $longestRestful = [int](($restfulPeriods | Measure-Object -Maximum).Maximum)
    $longestRestless = [int](($restlessPeriods | Measure-Object -Maximum).Maximum)

    $timeToSleep = [int]0
    if( $firstRestfulTime -gt $firstRestlessTime ){
        $timeToSleep = [int]$firstRestfulTime.Subtract($firstRestlessTime).TotalMinutes
    }

    [hashtable]$result = @{
        TimeToSleep = $timeToSleep
        LongestRestful = $longestRestful
        LongestRestless = $longestRestless
        LastRestlessRun = $lastRestlessRun
    }

    return $result
}


Export-ModuleMember -Function Start-SleepIqSession, Get-SleepIqBedInfo, Get-SleepIqSleepDayData, Get-SleepIqSleepSliceData, Import-SleepIqSleepSliceData
