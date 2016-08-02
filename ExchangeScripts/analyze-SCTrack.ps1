#analyze-SCTrack.ps1
# analyze a typical Symantec Cloud Tracking result csv

# vers: 3:53 PM 3/13/2014 - still blank $trkFile.Fullname... :S


param(
[string] $trkFile ="c:\usr\work\exch\scripts\SC-Track.csv"
)


#useful fields: fields: , Sender, Recipient, Delivery Status, Service Action, Service Reason	, SMTP Log Summary, Delivery Attempts, Sending Server Hostname, Sending Server IP, Message Id

# stock import, spac3es in field names (have to "quote" them to use them)
#$msgs= import-csv c:\usr\work\exch\scripts\SC-Track.csv | select "Sender","Recipient","Delivery Status","Service Action","Service Reason","SMTP Log Summary","Delivery Attempts","Sending Server Hostname","Sending Server IP","Message Id"

# drop spaces @{label="Latency(ms)";expression={($_.Latency.Milliseconds)}},
# @{label="SMTPStartDate";expression={($_."SMTP Start Date")}},
#$trkFile=gci "c:\usr\work\exch\scripts\SC-Track.csv" ;
$trkFile=gci $trkFile ;
$msgs= import-csv $trkFile | sort "SMTP Start Date" | select @{label="SMTPStartDate";expression={($_."SMTP Start Date")}},Sender,Recipient,@{label="DeliveryStatus";expression={($_."Delivery Status")}},@{label="ServiceAction";expression={($_."Service Action")}},@{label="ServiceReason";expression={($_."Service Reason")}},@{label="SMTPLogSummary";expression={($_."SMTP Log Summary")}},@{label="DeliveryAttempts";expression={($_."Delivery Attempts")}},@{label="SendingServerHostname";expression={($_."Sending Server Hostname")}},@{label="SendingServerIP";expression={($_."Sending Server IP")}},@{label="MessageId";expression={($_."Message Id")}} ;


# profile 
#$msgs= gc c:\usr\work\exch\scripts\SC-Track.csv | sort "SMTP Start Date"; $msgs | group "Delivery Status"

#$msgs= import-csv c:\usr\work\exch\scripts\SC-Track.csv | select "Sender","Recipient",@{label="DeliveryStatus";expression={($_."Delivery Status")}},@{label="ServiceAction";expression={($_."Service Action")}},@{label="ServiceReason";expression={($_."Service Reason")}},@{label="SMTPLogSummary";expression={($_."SMTP Log Summary")}},@{label="DeliveryAttempts";expression={($_."Delivery Attempts")}},@{label="SendingServerHostname";expression={($_."Sending Server Hostname")}},@{label="SendingServerIP";expression={($_."Sending Server IP")}},@{label="MessageId";expression={($_."Message Id")}} ; $msgs | group DeliveryStatus | select Count,Name | ft -auto | out-default; $msgs | group Recipient| select Count,Name | ft -auto | out-default ; $msgs | group Sender | select Count,Name | ft -auto | out-default;

# profile 
Write-Host -fore yellow ("`n" + (get-date).ToString("HH:mm:ss") + ":trkFile: " + $trkFile.FullName ) ;
Write-Host -fore green ("`nMsgs Count: " + $msgs.Count + "`nMsgs Range: '" + $msgs[0].SMTPStartDate + "' to '" + $msgs[-1].SMTPStartDate + "'") ;

Write-Host -fore yellow ("`n" + (get-date).ToString("HH:mm:ss") + ":Distribution of DeliveryStatus & ServiceReason:") ;
$msgs | group "DeliveryStatus",ServiceReason | select Count,Name | sort Name -desc | ft -auto | out-default ;

Write-Host -fore yellow ("`n" + (get-date).ToString("HH:mm:ss") + ":Distribution of Recipients:") ;
$msgs |  group Recipient| select Count,Name | sort Count -desc | ft -auto | out-default;

Write-Host -fore yellow ("`n" + (get-date).ToString("HH:mm:ss") + ":Distribution of Senders:") ;
$msgs | group Sender | select Count,Name| sort Count -desc |  ft -auto | out-default ;


