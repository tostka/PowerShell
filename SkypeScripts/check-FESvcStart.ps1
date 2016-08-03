#check-FESvcStart.ps1
# this just runs a 5second check loop on starting up the Front End Service 
<#

#distribute the cmd from tscclient share
$L13ProdALL="Server0;Server1;Server2;" ; $L13ProdALL=$L13ProdALL.split(";") ; $L13ProdALL | foreach { write-host -fore yell "copying c-scripts to $_" ; xcopy ("\\tsclient\c\usr\work\lync\scripts\check-FESvcStart.ps1") ("\\$_\c$\scripts\") /S /L; } ;

#>

$Svc="RTCSRV"; write-host –fore yell ( (get-date).ToString("HH:mm:ss") + ":Waiting for " + $Svc + " to start…") ; Do {write-host "." -NoNewLine; Start-Sleep -m (1000 * 5)} Until ((get-service RTCSRV).status -eq 'Running') ; write-host –fore yell ( (get-date).ToString("HH:mm:ss") + "Start Completed") ; 

