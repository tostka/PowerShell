# check-PingSweep.ps1
# run a subnet ping sweep on the designated subnet.

<#$tSubnet="192.168.1"
1..254 | foreach { 
  #"($tSubnet.$_)"
  $tAddr="($tSubnet.$_)" ; 
  $PingHash=@{quiet=$true ; count=1} ; 
  $Ping=(test-connection @PingHash -ComputerName "($tSubnet.$_)") ; 
  $hash=@{
      Address="$($tAddr)" ; 
      Ping = $Ping ;
  } ; 
  new-object psobject -prop $hash ; 
} # loop-E ; 
#>

1..254 | foreach { new-object psobject -prop @{Address="192.168.1.$_";Ping=(test-connection "192.168.1.$_" -quiet -count 1)}}
