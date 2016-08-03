#restart-CSServicesAll.ps1
# restarts all lync svcs and measures when FE svc starts

# this just runs a 5second check loop on starting up the Front End Service 

write-host –fore yell ( (get-date).ToString("HH:mm:ss") + ":Stopping Lync Svcs on " + $env:Computername ) ;
#stop-cswindowsservice -NoWait -Verbose ;#-Report $outransfile ; 
stop-cswindowsservice -Verbose ;#-Report $outransfile ; 

write-host –fore yell ( (get-date).ToString("HH:mm:ss") + ":Starting Lync Svcs on " + $env:Computername ) ;
# start all stopped
#get-cswindowservice | ?{$_.status -ne "Running"} | start-cswindowsservice -NoWait -Verbose -Report $outransfile ; 
# just start all
start-cswindowsservice -NoWait -Verbose -Report $outransfile ; 
$Svc="RTCSRV"; write-host –fore yell ( (get-date).ToString("HH:mm:ss") + ":Waiting for " + $Svc + " to start…") ; Do {write-host "." -NoNewLine; Start-Sleep -m (1000 * 5)} Until ((get-service RTCSRV).status -eq 'Running') ; write-host –fore yell ( (get-date).ToString("HH:mm:ss") + "Start Completed") ; 

