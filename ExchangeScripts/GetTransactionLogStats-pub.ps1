# GetTransactionLogStats.ps1

<# 
.SYNOPSIS
GetTransactionLogStats.ps1 - MS technet script to measure Ex2010/13 tlogging for upgrades to Ex13/16
.NOTES
Written By: Mike Hendrickson
Website:	https://blogs.technet.microsoft.com/exchange/2013/10/07/analyzing-exchange-transaction-log-generation-statistics/
Change Log
* 12:32 PM 8/4/2016 added pshelp, also added echo'd error when it doesn't find the TargetServers.txt:
* 6/22/2016 Version 2.1 most recent posted version as of 12:32 PM 8/4/2016
Example TargetServers.txt:
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Server0
SERVER2,DB1,DB2,DB3
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
1st line indicates to poll all db's on Server3 logs
# DOMAIN prod TargetServers.txt
#-=TargetServers.txt-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Server0
Server1
Server1
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Example schedtask cmd:
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
powershell.exe -noninteractive -noprofile -command "& {C:\LogStats\GetTransactionLogStats.ps1 -Gather -WorkingDirectory C:\LogStats}"
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# 
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
::: GetTransactionLogStats-ps1.cmd
powershell.exe -noprofile -noninteractive -command "& {D:\scripts\GetTransactionLogStats.ps1  -Gather -WorkingDirectory 'D:\Scripts\rpts\TLGS' -MonitoringExchange2013 $false}"
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# schedtask creation, hourly run at 5 after 1pm, wrapperless ps
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
schtasks /CREATE /TN GetTLS /TR "d:\scripts\GetTransactionLogStats-ps1.cmd" /sc hourly /st 13:05:00 /ru MYDOMAIN\user /rp [PASSWORD]
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
.DESCRIPTION
.PARAMETER  Gather
Switch specifying we want to capture current log generations. If this switch is omitted, the -Analyze switch must be used.
.PARAMETER  Analyze
Switch specifying we want to analyze already captured data. If this switch is omitted, the -Gather switch must be used.
.PARAMETER  ResetStats
Switch indicating that the output file, LogStats.csv, should be cleared and reset. Only works if combined with �Gather.
.PARAMETER  WorkingDirectory
The directory containing TargetServers.txt and LogStats.csv. If omitted, the working directory will be the current working directory of PowerShell (not necessarily the directory the script is in).
.PARAMETER  LogDirectoryOut
The directory to send the output log files from running in Analyze mode to. If omitted, logs will be sent to WorkingDirectory.
.PARAMETER  MaxSampleIntervalVariance
The maximum number of minutes that the duration between two samples can vary from 60. If we are past this amount, the sample will be discarded. Defaults to a value of 10.
.PARAMETER  MaxMinutesPastTheHour
How many minutes past the top of the hour a sample can be taken. Samples past this amount will be discarded. Defaults to a value of 15.
.PARAMETER  MonitoringExchange2013
Whether there are Exchange 2013/2016 servers configured in TargetServers.txt. Defaults to $true. If there are no 2013/2016 servers being monitored, set this to $false to increase performance.
.PARAMETER  DontAnalyzeInactiveDatabases
When running in Analyze mode, this specifies that any databases that have been found that did not generate any logs during the collection duration will be excluded from the analysis. This is useful in excluding passive databases from the analysis
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
Run typical DOMAIN DAG pass run from SITE650
.\GetTransactionLogStats.ps1 -Gather -WorkingDirectory "D:\Scripts\rpts\TLGS" -MonitoringExchange2013 $false
.EXAMPLE
Runs the script in Gather mode, taking a single snapshot of the current log generation of all configured databases:
PS C:\> .\GetTransactionLogStats.ps1 -Gather
.EXAMPLE
Runs the script in Gather mode, and indicates that no Exchange 2013/2016 servers are configured in TargetServers.txt:
PS C:\> .\GetTransactionLogStats.ps1 -Gather -MonitoringExchange2013 $false
.EXAMPLE
Runs the script in Gather mode, and changes the directory where TargetServers.txt is located, and where LogStats.csv will be written to:
PS C:\> .\GetTransactionLogStats.ps1 -Gather -WorkingDirectory "C:\GetTransactionLogStats" -ResetStats
.EXAMPLE
Runs the script in Analyze mode:
PS C:\> .\GetTransactionLogStats.ps1 -Analyze
.EXAMPLE
Runs the script in Analyze mode, and excludes database copies that did not generate any logs during the collection duration:
PS C:\> .\GetTransactionLogStats.ps1 -Analyze -DontAnalyzeInactiveDatabases $true
.EXAMPLE
Runs the script in Analyze mode, sending the output files for the analysis to a different directory. Specifies that only sample durations between 55-65 minutes are valid, and that each sample can be taken a maximum of 10 minutes past the hour before being discarded:
PS C:\> .\GetTransactionLogStats.ps1 -Analyze -LogDirectoryOut "C:\GetTransactionLogStats\LogsOut" -MaxSampleIntervalVariance 5 -MaxMinutesPastTheHour 10
.LINK
https://blogs.technet.microsoft.com/exchange/2013/10/07/analyzing-exchange-transaction-log-generation-statistics/
#>


#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages
#
#################################################################################

#Version 2.1
#Updated: 6/22/2016

[CmdletBinding()]
param(
	[switch] $Gather,
	[switch] $Analyze,
	[switch] $ResetStats,
	[string] $WorkingDirectory = "",
	[string] $LogDirectoryOut = "",
	[int]$MaxSampleIntervalVariance = 10,
	[int]$MaxMinutesPastTheHour = 15,
	[bool]$MonitoringExchange2013 = $true,
	[bool]$DontAnalyzeInactiveDatabases = $true
)

### GLOBAL VARIABLES ###
$gatherLogFileName = "LogStats.csv"
### END GLOBAL VARIABLES ###

#Function used to take a snapshot of the current log generation on all configured databases locally on the server.
#This function is intended to be used with Invoke-Command and executed on remote systems.
function GetLogGenerationsLocal
{
	[CmdletBinding()]
	[OutputType([PSObject[]])]
    param
    (
        [Bool]$MonitoringExchange2013,
        [string]$DatabaseList
    )

	#Return value contains an array of PSObjects containing log gen stats
    [PSObject[]]$logGenStats = @()

	#Check whether any specific databases were passed in
    if ([string]::IsNullOrEmpty($DatabaseList) -eq $false)
    {
        [string[]]$databases = $DatabaseList.Split(',')
    }
    else
    {
        [string[]]$databases = $null
    }

	#Get the local computer name for storing in the output variable
    $serverName = $env:COMPUTERNAME
                
    #Keep track of whether the server we're processing is 2013, since we'll have to use a different counter
    $is2013Server = $false
            
    #Keep track of all counters on the current server
    $allCounters = $null

    #Get the list of counters for the server. Try 2013 first, if configured.
    if ($MonitoringExchange2013 -eq $true)
    {
        $allCounters = Get-Counter -ListSet "MSExchangeIS HA Active Database" -ErrorAction SilentlyContinue
    }

    #Either we failed to connect to the server, or this isn't a 2013 server. Try the 2007/2010 command.
    if ($allCounters -eq $null)
    {
        $allCounters = Get-Counter -ListSet "MSExchange Database ==> Instances" -ErrorAction SilentlyContinue
    }
    else
    {
        $is2013Server = $true
    }
                
    #Got counters. Process them.
    if ($allCounters -ne $null)
    {
        #Set up the command to filter the counter list to the specific counters we want
        if ($is2013Server)
        {
            $targetCounterCommand = "`$allCounters.PathsWithInstances | where {`$_ -like '*Current Log Generation Number*' -and `$_ -notlike '*_total*'}"
        }
        else
        {
            $targetCounterCommand = "`$allCounters.PathsWithInstances | where {`$_ -like '*Information Store*Log File Current Generation*' -and `$_ -notlike '*Information Store/_Total*' -and `$_ -notlike '*Information Store/Base instance to*'}"
        }
                
        #DB's were specified for this server. Filter on them
        if ($databases -ne $null -and $databases.Count -gt 0)
        {
            $dbName = $databases[0].ToLower().Trim()
            $dbFilterString = " -and (`$_ -like '*$($dbName)*'"  
                    
            for ($i = 1; $i -lt $databases.Count; $i++)
            {
                $dbName = $databases[$i].ToLower().Trim()
                $dbFilterString += " -or `$_ -like '*$($dbName)*'"                                
            }
                    
            $targetCounterCommand = $targetCounterCommand.Replace("}", $dbFilterString + ")}")
        }

        #Invoke the command and get the counter names of databases we want
        $targetCounters = Invoke-Expression $targetCounterCommand

        #Process each counter in the list
        foreach ($counterName in $targetCounters)
        {
            #Parse out the database name from the current counter
            if ($is2013Server)
            {
                $dbNameStartIndex = $counterName.IndexOf("MSExchangeIS HA Active Database(") + "MSExchangeIS HA Active Database(".Length                            
            }
            else 
            {
                $dbNameStartIndex = $counterName.IndexOf("Instances(Information Store/") + "Instances(Information Store/".Length
            }
                        
            $dbNameEndIndex =  $counterName.IndexOf(")", $dbNameStartIndex)
            $dbName = $counterName.SubString($dbNameStartIndex, $dbNameEndIndex - $dbNameStartIndex)
                
            #Get the counter's value
            $counter = Get-Counter "$($counterName)" -ErrorAction SilentlyContinue
                        
            if ($counter -ne $null)
            {
                $logGenStat = New-Object PSObject

                $logGenStat | Add-Member NoteProperty Database $dbName
                $logGenStat | Add-Member NoteProperty Server $serverName
                $logGenStat | Add-Member NoteProperty LogGeneration $counter.CounterSamples[0].RawValue
                $logGenStat | Add-Member NoteProperty TimeCollected $([DateTime]::Now)

                $logGenStats += $logGenStat
            }
            else
            {
                Write-Error "ERROR: Failed to read perfmon counter from server $($serverName)"
            }
        }
    }
    else
    {
        Write-Error "ERROR: Failed to get perfmon counters from server $($serverName)"
    }
    
    return $logGenStats
}

#Function used to remotely initiate a local stat collection on all configured servers and databases.
function GetLogGenerationsFromRemoteServers
{
	[CmdletBinding()]
	param()

    #Read our input file of servers and databases
    $targetServersPath = AppendFileNameToDirectory -directory $WorkingDirectory -fileName "TargetServers.txt"
    
    [string[]]$targetServers = Get-Content -LiteralPath "$($targetServersPath)" -ErrorAction SilentlyContinue

    if ($targetServers -ne $null)
    {
        #Used to store all remote collection jobs that have been initiated
        [System.Management.Automation.Job[]]$allJobs = @()

        Write-Verbose "[$([DateTime]::Now)] Sending log generation collection job to $($targetServers.Count) servers."

        foreach ($server in $targetServers)
        {
            #Make sure we're not processing an empty line in the input file
            if ($server.Trim().Length -gt 0)
            {
                #Split the line into multiple columns. The first being the server name, the rest being database names.
                $serverName = $server.Split(',')[0].Trim()

                $wsManTest = $null
                $wsManTest = Test-WSMan -ComputerName $serverName -ErrorAction SilentlyContinue

                if ($wsManTest -eq $null)
                {
                    Write-Warning "[$([DateTime]::Now)] Failed to establish Remote PowerShell session to computer '$($serverName)'. To enable PowerShell Remoting, use Enable-PSRemoting, or 'winrm quickconfig' on the server to be configured. No databases on this server will be processed."
                    continue
                }

                $databaseList = ""

                if ($server.Contains(','))
                {
                    $databaseList = $server.Substring($server.IndexOf(",") + 1)
                }

                $job = Invoke-Command -ComputerName $serverName -ScriptBlock ${function:GetLogGenerationsLocal} -ArgumentList $MonitoringExchange2013,$databaseList -AsJob
                $allJobs += $job
            }
        }
        
        if ($allJobs.Count -gt 0)
        {          
            Write-Verbose "[$([DateTime]::Now)] Waiting for remote collections to finish."

            $silencer = Wait-Job $allJobs
            $results = Receive-Job $allJobs

            Write-Verbose "[$([DateTime]::Now)] Saving results to disk."

            $logPath = AppendFileNameToDirectory -directory $WorkingDirectory -fileName $gatherLogFileName
                    
            #The log file hasn't been created yet, or a reset was request, so add a header first
            if ($ResetStats -eq $true -or !(Test-Path -LiteralPath $logPath))
            {
                "Log Generation,Time Retrieved,ServerName,DatabaseName" | Out-File -FilePath $logPath -Encoding ASCII
            }

            foreach ($result in $results)
            {
                if ($result.GetType().Name -like "PSCustomObject")
                {
                    "$($result.LogGeneration),$($result.TimeCollected),$($result.Server),$($result.Database)" | Out-File -FilePath $logPath -Append -Encoding ASCII
                }
            }      

            Write-Verbose "[$([DateTime]::Now)] Finished saving results to disk."
        }
        else
        {
            Write-Error "[$([DateTime]::Now)] No servers in TargetServers.txt were reachable via Remote PowerShell. No logs will be processed."
        }
    } else {
        write-warning "$((get-date).ToString("HH:mm:ss")):WARNING! NO OR MISCONFIGGURED $($targetServersPath)!" ;
    }
}

#Function used to Analyze log files which were captured in Gather mode
function AnalyzeLogFiles
{
	[CmdletBinding()]
	param()

    #Get the UI culture for proper DateTime parsing (credit Thomas Stensitzki)
    $uiCulture = Get-UICulture

    $inputLogPath = AppendFileNameToDirectory -directory $WorkingDirectory -fileName $gatherLogFileName
    
    Write-Verbose "[$([DateTime]::Now)] Attempting to read: $($inputLogPath)"

    $inputCsv = Import-Csv -Path "$($inputLogPath)"

    if ($inputCsv -eq $null)
    {
        throw "Failed to read LogStats.csv at: $($inputLogPath)"
    }
    
    #Hash used to store the per database log generation readings. The Key will be a String. The Value will be a List of PSObjects
    $inputLogHash = @{}
    
    Write-Verbose "[$([DateTime]::Now)] Parsing Log Generation and Time Retrieved of $($inputCsv.Count) log entries."

    foreach ($line in $inputCsv)
    {
        $logEntry = New-Object PSObject @{"Log Generation"=[UInt64]::Parse($line."Log Generation");"Time Retrieved"=[DateTime]::Parse($line."Time Retrieved", $uiCulture);"ServerName"=$line.'ServerName';"DatabaseName"=$line.'DatabaseName'}

        $serverPlusDb = "$($line.'ServerName')-$($line.'DatabaseName')"

        #We haven't touched this database yet, so add it to the hashtable with a new List
        if (!($inputLogHash.ContainsKey($serverPlusDb)))
        {
            [PSObject[]]$logEntries = @()
            
            $inputLogHash.Add($serverPlusDb, $logEntries)
        }
        
        [PSObject[]]$logEntries = $inputLogHash[$serverPlusDb]
        $logEntries += $logEntry        
        $inputLogHash[$serverPlusDb] = $logEntries    
    }

    Write-Verbose "[$([DateTime]::Now)] Finished parsing Log Generation and Time Retrieved."

    #3 dimensional array to hold results.
    #1st dimension = Log number. Add an extra to total all databases
    #2nd dimension = Hour of day. Add an extra for all day totals
    #3rd dimension[0] is the total of logs created. [1] is the total of actual sample intervals. #[2] is the total number of samples
    $results = New-Object 'object[,,]' ($inputLogHash.Count + 1),25,3
    $logNames = New-Object 'object[]' ($inputLogHash.Count)

    #Keep track of any logs that we should skip outputting due to inactivity (if configured)
    [String[]]$logsToSkip = @()

    #Keep track of the log number since we can't index into a hashtable
    $logNum = -1
    
    #First loop through and process the data from the logs
    Write-Verbose "[$([DateTime]::Now)] Calculating log generation differences for log stats in $($inputLogHash.Keys.Count) unique server/database combinations."

    foreach ($kvp in $inputLogHash.GetEnumerator())
    {
        #First increment the log number
        $logNum++
        
        #Now get the log name and value
        $logNames[$logNum] = $kvp.Name
        $log = $kvp.Value                     

        #we need at least 2 lines to be able to compare samples
        if ($log.Count -ge 2)
        {
            if ($log[0].'Log Generation' -eq $log[$log.Count - 1].'Log Generation' -and $DontAnalyzeInactiveDatabases)
            {
                $logsToSkip += $kvp.Name
                continue
            }

            for ($j = 1; $j -lt $log.Count; $j++)
            {
                $logGenDifference = $log[$j].'Log Generation' - $log[$j - 1].'Log Generation'
                $timeSpan = New-TimeSpan -Start $log[$j - 1].'Time Retrieved' -End $log[$j].'Time Retrieved'
                
                #Only work on positive log differences
                #Only work on samples taken within maxMinutesPastTheHour minutes past the top of the hour
                #Only work on samples whose interval was within maxSampleIntervalVariance
                if ($logGenDifference -ge 0 -or $currentDate.Minute -le $MaxMinutesPastTheHour -or !($timeSpan.TotalMinutes -gt (60 + $MaxSampleIntervalVariance) -or $timeSpan.TotalMinutes -lt (60 - $MaxSampleIntervalVariance)))
                {
                    #Total this database for this hour
                    $results[$logNum, $log[$j - 1].'Time Retrieved'.Hour, 0] += $logGenDifference                    
                    $results[$logNum, $log[$j - 1].'Time Retrieved'.Hour, 1] += $timeSpan.TotalSeconds
                    $results[$logNum, $log[$j - 1].'Time Retrieved'.Hour, 2]++
                    
                    #Add this database totals to the entire days totals
                    $results[$logNum, 24, 0] += $logGenDifference                    
                    $results[$logNum, 24, 1] += $timeSpan.TotalSeconds
                    $results[$logNum, 24, 2]++
                    
                    #Add to all databases total
                    $results[($inputLogHash.Count), $log[$j - 1].'Time Retrieved'.Hour, 0] += $logGenDifference
                    $results[($inputLogHash.Count), $log[$j - 1].'Time Retrieved'.Hour, 1] += $timeSpan.TotalSeconds
                    $results[($inputLogHash.Count), $log[$j - 1].'Time Retrieved'.Hour, 2]++
                    
                    #Add to all databases totals for the day
                    $results[($inputLogHash.Count), 24, 0] += $logGenDifference
                    $results[($inputLogHash.Count), 24, 1] += $timeSpan.TotalSeconds
                    $results[($inputLogHash.Count), 24, 2]++
                }
                else
                {
                    continue
                }
            }
        }
        else
        {
            $logsToSkip += $kvp.Name
            continue
        }
    }

    if ($logsToSkip.Count -gt 0)
    {
        Write-Verbose "[$([DateTime]::Now)] Skipping $($logsToSkip.Count) database instances due to having less than 2 log generation samples."
    }

    Write-Verbose "[$([DateTime]::Now)] Calculating log generation averages per hour for $($inputLogHash.Keys.Count - $logsToSkip.Count) unique server/database combinations."

    [Hashtable]$analyzedFilesHT = @{}

    #Now output the results, and put together our averages for all servers
    for ($i = 0; $i -lt ($inputLogHash.Count + 1); $i++)
    {   
        if ($i -eq $inputLogHash.Count)
        {
            $logName = "AllDatabases"
        }
        else
        {            
            $logName = $logNames[$i]

            if ((StringArrayContainsString -targetArray $logsToSkip -strToCheck $logName) -eq $true)
            {
                continue
            }
        }
        
        $logPath = AppendFileNameToDirectory -directory $LogDirectoryOut -fileName "$($logName)-Analyzed.csv"
        [string[]]$logEntries = @()               
             
        $logsCreatedAllHours = $results[$i,24,0]
        $sampleIntervalSecondsAllHours = $results[$i,24,1]
        $numberSamplesAllHours = $results[$i,24,2]
        [Double]$totalRatioForCalc = 0 #Keeps track of the sum of each hour's PercentDailyUsageForCalc for this log

        for ($j = 0; $j -lt 24; $j++)
        {
            $logsCreatedThisHour = $results[$i,$j,0]
            $sampleIntervalSecondsThisHour = $results[$i,$j,1]
            $numberSamplesThisHour = $results[$i,$j,2]

            if ($logsCreatedThisHour -ne $null -and $sampleIntervalSecondsThisHour -ne $null -and $numberSamplesThisHour -ne $null)
            {
                $averageSamplePer60Minutes = $logsCreatedThisHour / $sampleIntervalSecondsThisHour * 3600               
                $averageSampleSizeForHour = $logsCreatedThisHour / $numberSamplesThisHour

                if ($logsCreatedAllHours -gt 0)
                {
                    $thisHourToAllLogRatio = $logsCreatedThisHour / $logsCreatedAllHours                   
                }
                else
                {
                    $thisHourToAllLogRatio = 0                   
                }

                $thisHourToAllLogPercent = $thisHourToAllLogRatio * 100

                $totalRatioForCalc += $thisHourToAllLogRatio
                
                if ($j -eq 23) #If this is the last hour of the day, fill in the remaining space so that all hours add up to 100% in the PercentDailyUsageForCalc column
                {
                    $remainderPercentage = 1 - $totalRatioForCalc
                        
                    if ($remainderPercentage -gt 0)
                    {
                        if ($remainderPercentage -gt .001) #Only bother informing of adding the remainder if the remainder percent is signficant. Arbitrarily choosing .001 for significance.
                        {
                            Write-Verbose "[$([DateTime]::Now)] Adding $($remainderPercentage) to hour 23 for PercentDailyUsageForCalc column of instance $($logName) so that the total of rows equals 100%."
                        }

                        $thisHourToAllLogRatio += $remainderPercentage
                    }
                }                           
                
                $logEntries += "$($j),$($logsCreatedThisHour),$($sampleIntervalSecondsThisHour),$($numberSamplesThisHour),$($averageSampleSizeForHour),$($thisHourToAllLogPercent),$($thisHourToAllLogRatio),$($averageSamplePer60Minutes)"
            }
        }

        $analyzedFilesHT.Add($logPath, $logEntries)
    }

    Write-Verbose "[$([DateTime]::Now)] Saving results to disk."

    foreach ($file in $analyzedFilesHT.Keys)
    {
        "Hour,TotalLogsCreated,TotalSampleIntervalSeconds,NumberOfSamples,AverageSample,PercentDailyUsage,PercentDailyUsageForCalc,AverageSamplePer60Minutes" | Out-File -FilePath $file -Encoding ASCII

        [string[]]$logEntries = $analyzedFilesHT[$file]

        foreach ($entry in $logEntries)
        {
            $entry | Out-File -FilePath $file -Append -Encoding ASCII
        }
    }

    #Finally, get the heat map
    $heatMapPath = AppendFileNameToDirectory -directory $LogDirectoryOut -fileName "HeatMap-AllCopies.csv"
    $heatMapCombinedPath = AppendFileNameToDirectory -directory $LogDirectoryOut -fileName "HeatMap-DBsCombined.csv"

    $heatMapDBsCombined = @{}

    "ServerName,DatabaseName,LogsGenerated,DurationInSeconds" | Out-File -FilePath $heatMapPath -Encoding ASCII
    "DatabaseName,LogsGenerated" | Out-File -FilePath $heatMapCombinedPath -Encoding ASCII
    
    Write-Verbose "[$([DateTime]::Now)] Generating database heat map."

    foreach ($kvp in $inputLogHash.GetEnumerator())
    {
        if ($kvp.Value.Count -gt 1)
        {
            $sortedByDate = $kvp.Value | Sort-Object -Property "Time Retrieved"

            $logDifference = $kvp.Value[$kvp.Value.Count - 1].'Log Generation' - $kvp.Value[0].'Log Generation'

            if ($logDifference -eq 0 -and $DontAnalyzeInactiveDatabases -eq $true)
            {
                continue
            }

            $timeSpan = New-TimeSpan -Start $kvp.Value[0].'Time Retrieved' -End $kvp.Value[$kvp.Value.Count - 1].'Time Retrieved'

            if ($heatMapDBsCombined.ContainsKey($kvp.Value[0].'DatabaseName') -eq $false)
            {
                $heatMapDBsCombined.Add($kvp.Value[0].'DatabaseName', $logDifference)
            }
            else
            {
                $value = $heatMapDBsCombined[$kvp.Value[0].'DatabaseName']
                $value += $logDifference
                $heatMapDBsCombined[$kvp.Value[0].'DatabaseName'] = $value
            }
            
            "$($kvp.Value[0].'ServerName'),$($kvp.Value[0].'DatabaseName'),$($logDifference),$($timeSpan.TotalSeconds)" | Out-File -FilePath $heatMapPath -Append -Encoding ASCII
        }
    }

    #Output the combined heat map
    foreach ($kvp in $heatMapDBsCombined.GetEnumerator())
    {
        "$($kvp.Key),$($kvp.Value)" | Out-File -FilePath $heatMapCombinedPath -Append -Encoding ASCII
    }

    Write-Verbose "[$([DateTime]::Now)] Finished analying log stats."
}

#Used to strip the slash at the end of a file path
function StripTrailingSlash
{
	[CmdletBinding()]
    param($stringIn)
    
    if ($stringIn.EndsWith("\") -or $stringIn.EndsWith("/"))
    {
        $stringIn = $stringIn.Substring(0, ($stringIn.Length - 1))
    }
    
    return $stringIn
}

#Returns a full path (relative or absolute) consisting of the given directory plus the given filename
function AppendFileNameToDirectory
{
	[CmdletBinding()]
    param($directory, $fileName)
    
    if ($directory -eq "")
    {
        return $fileName
    }
    else
    {
        return "$($directory)\$($fileName)"
    }
}

function StringArrayContainsString
{
	[CmdletBinding()]
    param([String[]]$targetArray, [String]$strToCheck)
    
    [bool]$containsString = $false
    
    if ($targetArray -ne $null -and $targetArray.Count -gt 0)
    {
        foreach ($str in $targetArray)
        {
            if ($str -like $strToCheck)
            {
                $containsString = $true
                break
            }
        }
    }
    
    return $containsString
}

# Function that returns true if the incoming argument is a help request
function IsHelpRequest
{
	[CmdletBinding()]
	param($argument)

	return ($argument -eq "-?" -or $argument -eq "-help");
}

# Function that displays the help related to this script following
# the same format provided by get-help or <cmdletcall> -?
function Usage
{
	[CmdletBinding()]
	param()
@"

NAME:
`tGetTransactionLogStats.ps1

SYNOPSIS:
`tUsed to collect and analyze Exchange transaction log generation statistics.
`tDesigned to be run as an hourly scheduled task, on the top of each hour.
`tCan be run against one or more servers and databases.

SYNTAX:
`tGetTransactionLogStats.ps1
`t`t[-Gather]
`t`t[-Analyze]
`t`t[-ResetStats]
`t`t[-WorkingDirectory <StringValue>]
`t`t[-LogDirectoryOut <StringValue>]
`t`t[-MaxSampleIntervalVariance <IntegerValue>]
`t`t[-MaxMinutesPastTheHour <IntegerValue>]
`t`t[-MonitoringExchange2013 <BooleanValue>]
`t`t[-DontAnalyzeInactiveDatabases <BooleanValue>]

PARAMETERS:
`t-Gather
`t`tSwitch specifying we want to capture current log generations.
`t`tIf this switch is omitted, the -Analyze switch must be used.

`t-Analyze
`t`tSwitch specifying we want to analyze already captured data.
`t`tIf this switch is omitted, the -Gather switch must be used.

`t-ResetStats
`t`tSwitch indicating that the output file, LogStats.csv, should
`t`tbe cleared and reset. Only works if combined with �Gather.

`t-WorkingDirectory
`t`tThe directory containing TargetServers.txt and LogStats.csv.
`t`tIf omitted, the working directory will be the current working
`t`tdirectory of PowerShell (not necessarily the directory the
`t`tscript is in).

`t-LogDirectoryOut
`t`tThe directory to send the output log files from running in
`t`tAnalyze mode to. If omitted, logs will be sent to WorkingDirectory.

`t-MaxSampleIntervalVariance
`t`tThe maximum number of minutes that duraction between two samples can
`t`tvary from 60. If we are past this amount, the sample will be discarded.
`t`t.Defaults to a value of 10.

`t-MaxMinutesPastTheHour
`t`tHow many minutes past the top of the hour a sample can be taken.
`t`tSamples past this amount will be discarded. Defaults to a value of 15.

`t-MonitoringExchange2013
`t`tWhether there are Exchange 2013 servers configured in TargetServers.txt.
`t`tDefaults to `$true. If there are no 2013 servers being monitored, set this
`t`tto `$false to increase performance.

`t-DontAnalyzeInactiveDatabases
`t`tWhen running in Analyze mode, this specifies that any databases
`t`tthat have been found that did not generate any logs during the collection
`t`tduration will be excluded from the analysis. This is useful in excluding
`t`tpassive databases from the analysis.


`t-------------------------- EXAMPLES ----------------------------

PS C:\> .\GetTransactionLogStats.ps1 -Gather

PS C:\> .\GetTransactionLogStats.ps1 -Gather -MonitoringExchange2013 `$false

PS C:\> .\GetTransactionLogStats.ps1 -Gather -WorkingDirectory "C:\GetTransactionLogStats" -ResetStats

PS C:\> .\GetTransactionLogStats.ps1 -Analyze

PS C:\> .\GetTransactionLogStats.ps1 -Analyze -DontAnalyzeInactiveDatabases `$true

PS C:\> .\GetTransactionLogStats.ps1 -Analyze -LogDirectoryOut "C:\GetTransactionLogStats\LogsOut" -MaxSampleIntervalVariance 5 -MaxMinutesPastTheHour 10
"@
}

####################################################################################################
# Script starts here
####################################################################################################

# Check for Usage Statement Request
$args | foreach { if (IsHelpRequest $_) { Usage; exit; } }

#Do input validation before proceeding
if (($Gather -eq $false -and $Analyze -eq $false) -or ($Gather -eq $true -and $Analyze -eq $true))
{
    Write-Host -ForeGroundColor Red "ERROR: Either the Gather or Analyze switch must be specified, but not both."
}
elseif ($WorkingDirectory -ne "" -and !(Test-Path -LiteralPath $WorkingDirectory))
{
    Write-Host -ForeGroundColor Red "ERROR: Working directory '$($WorkingDirectory)' must be created before proceeding."
}
elseif ($Analyze -eq $true -and $LogDirectoryOut -ne "" -and !(Test-Path -LiteralPath $LogDirectoryOut))
{
    Write-Host -ForeGroundColor Red "ERROR: Output log directory '$($LogDirectoryOut)' must be created before proceeding."
}
elseif ($Analyze -eq $true -and ($MaxSampleIntervalVariance -lt 0 -or $MaxMinutesPastTheHour -lt 0))
{
    Write-Host -ForeGroundColor Red "ERROR: MaxSampleIntervalVariance and MaxMinutesPastTheHour must have non-negative values."
}
else #Made it past input validation
{
    #Massage the log directory string so they're in an expected format when we need them
    $LogDirectoryOut = StripTrailingSlash($LogDirectoryOut)
    $WorkingDirectory = StripTrailingSlash($WorkingDirectory)
     
    #Now do the real work
    if ($Gather -eq $true)
    {
        GetLogGenerationsFromRemoteServers
    }
    else #(Analyze -eq $true)
    {
        AnalyzeLogFiles
    }
}
