# clear-sessions.ps1
# 1:03 PM 2/11/2016 initial vers

#Get-PSSession|?{$_.State -eq 'Broken'} | Remove-PSSession #-whatif
if($psbs=Get-PSSession|?{$_.State -eq 'Broken'} ){
foreach ($psb in $psbs) {
  "removing broken session: $($psb.id)`t$($psb.name)`t$($psb.state)`t$($psb.ConfigurationName)" ;
  Remove-PSSession $psb.id -whatif
} ; 
} else {
  "(no broken sessions found)" ;
} ; 