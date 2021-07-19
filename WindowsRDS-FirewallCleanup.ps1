#Script derived from user js2010's Script
#Source https://stackoverflow.com/a/40915201

$profiles = get-wmiobject -class win32_userprofile
$targetos = 0

## Interactive Menu
function Show-Menu
{
    param (
        [string]$Title = 'Select Windows Server version'
    )
    Clear-Host
    Write-Host $Title -ForegroundColor Green
    
    Write-Host "1: Windows Server 2019"
    Write-Host "2: Windows Server 2016"
    Write-Host "Q: Quit"

    return Read-Host "Selection : "
}

#Catch argument / -2019 OR -2016
if ($args.count -gt 0){ 
  $param1=$args[0]
  switch ($param1)
    {
        '-2019' {
            $targetos = '1'
            }
        '-2016' {
            $targetos = '2'
            }
        default {
            $targetos = Show-Menu
            } 
    }
}

# OS Detection
if ($targetos -eq 0) {
  $osversion = Get-CimInstance Win32_Operatingsystem | Select-Object -expand Caption
  if ($osversion -Match 'Windows Server 2019'){
    Write-Host "WINDOWS SERVER 2019 DETECTED"
    $targetos = 1
  }
  elseif ($osversion -Match 'Windows Server 2016') {
    Write-Host "WINDOWS SERVER 2019 DETECTED" 
    $targetos = 2
  } 
  else {
    Write-Host "OS NOT DETECTED"
    $targetos = Show-Menu
  }
}

Clear-Host
# cleanupPath variable assignation
switch ($targetos)
 {
  '1' {
      Write-Host "Windows Server 2019 Selected"
      $cleanupPath = "HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\RestrictedServices\AppIso\FirewallRules"
  } 
  '2' {
      Write-Host "Windows Server 2016 Selected"
      $cleanupPath = "HKLM:\System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\RestrictedServices\Configurable\System"
  } 
  'q' {
      return
  }
  default {
    return
    }
 }

# Registry patch
New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy' -Name  'DeleteUserAppContainersOnLogoff' -Value '1' -PropertyType 'DWORD' –Force
Write-Host "Registry patched !" -ForegroundColor Green

#Start cleanup script
Write-Host "Getting Firewall Rules" -ForegroundColor Yellow

# deleting rules with no owner would be disastrous
$Rules = Get-NetFirewallRule -All | 
  Where-Object {$profiles.sid -notcontains $_.owner -and $_.owner }

Write-Host "Getting Firewall Rules from ConfigurableServiceStore Store" -ForegroundColor Yellow

$rules2 = Get-NetFirewallRule -All -PolicyStore ConfigurableServiceStore | 
  Where-Object { $profiles.sid -notcontains $_.owner -and $_.owner }

$total = $rules.count + $rules2.count
Write-Host ">> " $total "Firewall Rules to delete . . ." -ForegroundColor Green

$result = measure-command {

  # tracking
  $start = Get-Date; $i = 0.0 ; 
  
  foreach($rule in $rules){

    # action
    remove-itemproperty -path "HKLM:\System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules" -name $rule.name

    # progress
    $i = $i + 1.0
    $prct = $i / $total * 100.0
    $elapsed = (Get-Date) - $start; 
    $totaltime = ($elapsed.TotalSeconds) / ($prct / 100.0)
    $remain = $totaltime - $elapsed.TotalSeconds
    $eta = (Get-Date).AddSeconds($remain)

    # display
    $prctnice = [math]::round($prct,2) 
    $elapsednice = $([string]::Format("{0:d2}:{1:d2}:{2:d2}", $elapsed.hours, $elapsed.minutes, $elapsed.seconds))
    $speed = $i/$elapsed.totalminutes
    $speednice = [math]::round($speed,2) 
    Write-Progress -Activity "Deleting rules ETA $eta elapsed $elapsednice loops/min $speednice" -Status "$prctnice" -PercentComplete $prct -secondsremaining $remain
  }


  # tracking
  # $start = Get-Date; $i = 0 ; $total = $rules2.Count

  foreach($rule2 in $rules2) {
        
    remove-itemproperty -path $cleanupPath -name $rule2.name

    # progress
    $i = $i + 1.0
    $prct = $i / $total * 100.0
    $elapse = (Get-Date) - $start; 
    $totaltime = ($elapsed.TotalSeconds) / ($prct / 100.0)
    $remain = $totaltime - $elapsed.TotalSeconds
    $eta = (Get-Date).AddSeconds($remain)

    # display
    $prctnice = [math]::round($prct,2) 
    $elapsednice = $([string]::Format("{0:d2}:{1:d2}:{2:d2}", $elapsed.hours, $elapsed.minutes, $elapsed.seconds))
    $speed = $i/$elapsed.totalminutes
    $speednice = [math]::round($speed,2) 
    Write-Progress -Activity "Deleting rules2 ETA $eta elapsed $elapsednice loops/min $speednice" -Status "$prctnice" -PercentComplete $prct -secondsremaining $remain
  }
}

$end = get-date
write-host 'END - ' $end 
write-host 'ETA - ' $eta

write-host $result.minutes min $result.seconds sec