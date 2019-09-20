$startTime = Get-Date
$CurrentDate = Get-Date -format F
$myFQDN = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain

$strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)

$CrawlDuration = @()
$CrawlRate = @()

$SuccessCol = "#C6EFCE"
$ErrorCol   = "#FFC7CE"
$WarningCol = "#FFEB9C"

$tempfolder = "d:\temp\"


Add-PSSnapin Microsoft.SharePoint.Powershell -ErrorAction SilentlyContinue

#Get CA URL
$CAURL = Get-SPWebApplication -includecentraladministration | where {$_.IsAdministrationWebApplication} | Select-Object -ExpandProperty Url
 
#Get the search service application
$SSA = Get-SPEnterpriseSearchServiceApplication #-Identity "Search Service Application Name"
 
#Get all content sources
$ContentSources = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $SSA #| where {$_.Name -eq $ContentSourceName}

$ResultsCount = 1000
$ReportDate = Get-Date -format "dd-MM-yyyy"
$Col = "#ddd"

$CrawlLogURL = $CAURL +"_admin/search/CrawlLogCrawls.aspx?appid={$($SSA.Id)" +"}"
$CrawlErrorURL = $CAURL +"_admin/search/CrawlLogErrors.aspx?appid={$($SSA.Id)" +"}"
$AdministrationURL = $CAURL +"searchadministration.aspx?appid={$($SSA.Id)" +"}"
$SPHAURL = $CAURL +"/Lists/HealthReports/AllItems.aspx"
$SPServiceAppURL = $CAURL +"_admin/ServiceApplications.aspx"
$SPServicesOnServerURL = $CAURL +"_admin/Server.aspx"


$emailBody = @"
<html>
 <body style='font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
  <h2>SharePoint Health Check Report<br>$CurrentDate</h2>
  <br />
"@

$consol = @()

$webapplications = Get-SPWebApplication -includecentraladministration
foreach ($webapplication in $webapplications) {
    $object = New-Object -TypeName PSObject
    $object | Add-Member -Name 'Name' -MemberType Noteproperty -Value $webapplication.ApplicationPool.Name
    $object | Add-Member -Name 'Status' -MemberType Noteproperty -Value $webapplication.ApplicationPool.Status
    $object | Add-Member -Name 'Username' -MemberType Noteproperty -Value $webapplication.ApplicationPool.Username
    $consol += $object
}



$serviceapps = Get-SPServiceApplicationPool
foreach ($serviceapp in $serviceapps) {
    $object = New-Object -TypeName PSObject
    $object | Add-Member -Name 'Name' -MemberType Noteproperty -Value $serviceapp.Name
    $object | Add-Member -Name 'Status' -MemberType Noteproperty -Value $serviceapp.Status
    $object | Add-Member -Name 'Username' -MemberType Noteproperty -Value $serviceapp.ProcessAccountName
    $consol += $object
}

$consol = $consol | select-object Status, Username, Name -Unique

$spservers = Get-SPServer | where {$_.Role -ne "Invalid"}
$option = New-PSSessionOption -ProxyAccessType NoProxyServer

$IISAppPools = @()

foreach ($spserver in $spservers) {
    
    $scriptblock = {
                Import-Module WebAdministration
                foreach ($AppPool in $using:consol) {
                    $output = Get-IISAppPool -Name $AppPool.Name
                    $output
                   }
    }
    $appPoolStatus = Invoke-Command $scriptblock -ComputerName $spserver.Address -SessionOption $option
    $IISAppPools += $appPoolStatus
}

#$IISAppPools
$Consolfinal = @()

foreach ($IISAppPool in $IISAppPools) {
    $object = New-Object -TypeName PSObject
    $object | Add-Member -Name 'Name' -MemberType Noteproperty -Value $IISAppPool.Name
    $object | Add-Member -Name 'IISStatus' -MemberType Noteproperty -Value $IISAppPool.State
    $object | Add-Member -Name 'Server' -MemberType Noteproperty -Value $IISAppPool.PSComputerName
    foreach ($SPAppPool in $consol) {
        if ($SPAppPool.Name -eq $IISAppPool.Name) {
            $object | Add-Member -Name 'SPStatus' -MemberType Noteproperty -Value $SPAppPool.Status
            $object | Add-Member -Name 'Username' -MemberType Noteproperty -Value $SPAppPool.Username
        }
    }
    $Consolfinal += $object
}

#$Consolfinal | ft

$UniqueAppPools = $Consolfinal | select-object SPStatus, Username, Name -Unique | sort-object Name
$UniqueServers = $Consolfinal | select-object Server -Unique

$NewAppPools = @()

foreach ($UniqueAppPool in $UniqueAppPools) {
    $object = New-Object -TypeName PSObject
    $object | Add-Member -Name 'Name' -MemberType Noteproperty -Value $UniqueAppPool.Name
    $object | Add-Member -Name 'SPStatus' -MemberType Noteproperty -Value $UniqueAppPool.SPStatus
    $object | Add-Member -Name 'Username' -MemberType Noteproperty -Value $UniqueAppPool.Username

    foreach ($NewAppPool in $Consolfinal) {
        if ($NewAppPool.Name -eq $UniqueAppPool.Name) {
            $object | Add-Member -Name $NewAppPool.Server -MemberType Noteproperty -Value $NewAppPool.IISStatus
        }

    }
    $NewAppPools += $object

}

$emailBody += @"
  <h2>IIS App Pool Status</h2>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
    <tr bgcolor="#3366cc" style="color:#ffffff">
    <th style="padding:5px">Application Pool</th>
    <th style="padding:5px">SharePoint Status</th>
    <th style="padding:5px">Service Account</th>
"@

foreach ($Server in $UniqueServers) {

$emailBody += @"  
    <th style="padding:5px">$($Server.Server)</th>
    
"@   
}

$emailBody += @"
   </tr>
"@

$Col = "#ddd"

$emailbody += @"
                
"@

                foreach ($NewAppPool in $NewAppPools) {
                    if($Col -eq "#fff") {
                        $Col = "#ddd"
                    }
                    else {
                        $Col = "#fff"
                    }
                    $emailBody += @"  
                    <tr style='background-color:$Col'>
                    <td style="padding:5px">$($NewAppPool.Name)</th>
"@ 
                    if ($NewAppPool.SPStatus -eq "Online") {
                        $emailBody += @"
                            <td style="padding:5px;background-color:$SuccessCol">$($NewAppPool.SPStatus)</th>
"@ 
                    } else {
                        $emailBody += @"
                            <td style="padding:5px;background-color:$ErrorCol">$($NewAppPool.SPStatus)</th>
"@ 
                    }
                    $emailBody += @"
                    <td style="padding:5px">$($NewAppPool.Username)</th>

"@   
                    foreach ($Server in $UniqueServers) {
                        if ($NewAppPool.($Server.Server) -eq "Started") {
                            $emailBody += @"  
                            <td style="padding:5px;background-color:$SuccessCol">$($NewAppPool.($Server.Server))</th>

"@   
                        } elseif (($NewAppPool.($Server.Server) -eq $null) -or ($NewAppPool.($Server.Server) -eq "")) {
                            $emailBody += @"  
                            <td style="padding:5px">$($NewAppPool.($Server.Server))</th>

"@   
                        } else {
                            $emailBody += @"  
                            <td style="padding:5px;background-color:$ErrorCol">$($NewAppPool.($Server.Server))</th>

"@
                        }
                    }
                    $emailbody += @"
                     </tr>
"@
                }


$emailbody += @"
                    </table>
                    <br />
"@


$FinalConsolIISSites = @()

foreach ($spserver in $spservers) {
    
    $scriptblock = {
        $webapplications = $args     
        Import-Module -Name WebAdministration

        $bindings = Get-WebBinding | select protocol, bindingInformation, certificateHash, @{name='ItemXPath';expression= {$_.ItemXPath -replace "(?:.*?)name='|' and.*" ,''}}


        $websites = @()


        foreach ($webapplication in $webapplications) {

            $websites += Get-Website -Name $webapplication

        }
        $websites += Get-Website -Name "SharePoint Web Services"
        #$websites | select Name, State

        $Certs = @()

        $Certs = Get-ChildItem "Cert:\LocalMachine\My" | select DnsNameList, NotAfter,HasPrivateKey, Thumbprint, @{name='CN';expression= {$_.Subject -replace "(?:.*?)CN=|, .*" ,''}}
        $Certs += Get-ChildItem "Cert:\LocalMachine\SharePoint" | select DnsNameList, NotAfter,HasPrivateKey, Thumbprint, @{name='CN';expression= {$_.Subject -replace "(?:.*?)CN=|, .*" ,''}}


        $ConsolIISSites =  @()

        foreach ($binding in $bindings) {  
            if ($websites.Name.Contains($binding.ItemXPath)) {
                $object = New-Object -TypeName PSObject
                $object | Add-Member -Name 'Server' -MemberType Noteproperty -Value (Get-WmiObject win32_computersystem).DNSHostName
                foreach ($Website in $websites) {
                    if ($Website.Name -eq $binding.ItemXPath) {
                        $object | Add-Member -Name 'WebSiteName' -MemberType Noteproperty -Value $Website.Name
                        $object | Add-Member -Name 'WebSiteState' -MemberType Noteproperty -Value $Website.State
                    }
                }
                $object | Add-Member -Name 'Protocol' -MemberType Noteproperty -Value $binding.protocol
                $object | Add-Member -Name 'BindingInfo' -MemberType Noteproperty -Value $binding.bindingInformation
                if ($binding.certificateHash) {
                    $object | Add-Member -Name 'Thumbprint' -MemberType Noteproperty -Value $binding.certificateHash
                    foreach ($Cert in $Certs) {
                        if ($Cert.Thumbprint -eq $binding.certificateHash) {
                            $object | Add-Member -Name 'CertExpiry' -MemberType Noteproperty -Value $Cert.NotAfter
                            $object | Add-Member -Name 'CN' -MemberType Noteproperty -Value $Cert.CN
                            $object | Add-Member -Name 'PrivateKey' -MemberType Noteproperty -Value $Cert.HasPrivateKey
                            $object | Add-Member -Name 'SANS' -MemberType Noteproperty -Value $Cert.DnsNameList
                        }
                    }
                } else {
                    $object | Add-Member -Name 'Thumbprint' -MemberType Noteproperty -Value "None"
                    $object | Add-Member -Name 'CertExpiry' -MemberType Noteproperty -Value "n/a"
                    $object | Add-Member -Name 'CN' -MemberType Noteproperty -Value "n/a"
                    $object | Add-Member -Name 'PrivateKey' -MemberType Noteproperty -Value "n/a"
                    $object | Add-Member -Name 'SANS' -MemberType Noteproperty -Value "n/a"
                }
                $ConsolIISSites += $object 
            }
        }
        $ConsolIISSites

}
    $ConsolIISSites = Invoke-Command $scriptblock -ComputerName $spserver.Address -SessionOption $option -ArgumentList $webapplications.DisplayName
    $FinalConsolIISSites += $ConsolIISSites
}



$emailBody += @"
  <h2>IIS Web Site Status</h2>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
    <tr bgcolor="#3366cc" style="color:#ffffff">
    <th style="padding:5px">Server</th>
    <th style="padding:5px">IIS Web Site</th>
    <th style="padding:5px">Status</th>
    <th style="padding:5px">Protocol</th>
    <th style="padding:5px">Binding</th>
    <th style="padding:5px">Certificate Thumbprint</th>
    <th style="padding:5px">Certificate Expiry</th>
    <th style="padding:5px">Certificate CN</th>
    <th style="padding:5px">Private Key</th>
    <th style="padding:5px">SANS</th>
    </tr>
"@

$Col = "#ddd"

foreach ($FinalConsolIISSite in $FinalConsolIISSites) {
            if($Col -eq "#fff") {
                $Col = "#ddd"
            }
            else {
                $Col = "#fff"
            }
            $emailbody += @"
                <tr style='background-color:$Col'>
                <td style="padding:5px">$($FinalConsolIISSite.Server)</td>
                <td style="padding:5px">$($FinalConsolIISSite.WebSiteName)</td>
"@
            if ($FinalConsolIISSite.WebSiteState -eq "Started") {
                $emailbody += @"
                    <td style="padding:5px;background-color:$SuccessCol">$($FinalConsolIISSite.WebSiteState)</td>
"@
            } else {
                $emailbody += @"
                    <td style="padding:5px;background-color:$ErrorCol">$($FinalConsolIISSite.WebSiteState)</td>
"@
            }
            $emailbody += @"
                <td style="padding:5px">$($FinalConsolIISSite.Protocol)</td>
                <td style="padding:5px">$($FinalConsolIISSite.BindingInfo)</td>
"@
            if (($FinalConsolIISSite.Protocol -eq "https") -and ($FinalConsolIISSite.Thumbprint -eq "None")) {
                $emailbody += @"
                    <td style="padding:5px;background-color:$ErrorCol">$($FinalConsolIISSite.Thumbprint)</td>
"@
            } else {
                $emailbody += @"
                    <td style="padding:5px">$($FinalConsolIISSite.Thumbprint)</td>
"@
            }
            if ($FinalConsolIISSite.CertExpiry -eq "n/a") {
            $emailbody += @"
                <td style="padding:5px">$($FinalConsolIISSite.CertExpiry)</td>
"@
            } elseif (($FinalConsolIISSite.CertExpiry -lt $startTime.AddDays(30)) -AND ($FinalConsolIISSite.CertExpiry -gt $startTime.AddDays(14))) {
            $emailbody += @"
                <td style="padding:5px;background-color:$WarningCol">$(get-date $FinalConsolIISSite.CertExpiry -format F)</td>
"@
            } elseif ($FinalConsolIISSite.CertExpiry -gt $startTime.AddDays(30) ) {
            $emailbody += @"
                <td style="padding:5px;background-color:$SuccessCol">$(get-date $FinalConsolIISSite.CertExpiry -format F)</td>
"@
            } else {
            $emailbody += @"
                <td style="padding:5px;background-color:$ErrorCol">$(get-date $FinalConsolIISSite.CertExpiry -format F)</td>
"@
            }
            $emailbody += @"
                <td style="padding:5px">$($FinalConsolIISSite.CN)</td>
                <td style="padding:5px">$($FinalConsolIISSite.PrivateKey)</td>
                <td style="padding:5px">$($FinalConsolIISSite.SANS)</td>
"@     
            $emailbody += @"
                </tr>

"@          

}


$emailbody += @"
                    </table>
                    <br />
"@


$ServiceApps = Get-SPServiceApplication | sort-object DisplayName
$ServiceApps += Get-SPServiceApplicationProxy | sort-object DisplayName
$ServiceApps = $ServiceApps

$emailBody += @"
  <h2>Service Application and Proxy Status</h2>
  <p style='font-size:8pt;'>$SPServiceAppURL</p>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
    <tr bgcolor="#3366cc" style="color:#ffffff">
    <th style="padding:5px">Service Application</th>
    <th style="padding:5px">Type</th>
    <th style="padding:5px">Status</th>
    </tr>
"@

$Col = "#ddd"

foreach ($ServiceApp in $ServiceApps) {
            if($Col -eq "#fff") {
                $Col = "#ddd"
            }
            else {
                $Col = "#fff"
            }
            $emailbody += @"
                <tr style='background-color:$Col'>
                <td style="padding:5px">$($ServiceApp.DisplayName)</td>
                <td style="padding:5px">$($ServiceApp.TypeName)</td>
"@     
            if ($ServiceApp.Status -eq "Online") {
                $emailbody += @"                
                    <td style="padding:5px;background-color:$SuccessCol">$($ServiceApp.Status)</td>
"@
            } elseif (($ServiceApp.Status -eq "Disabled") -and ($ServiceApp.TypeName -eq "Usage and Health Data Collection Proxy")) {
                $emailbody += @"                
                    <td style="padding:5px;background-color:$WarningCol">$($ServiceApp.Status)</td>
"@
            } else {
                $emailbody += @"                
                    <td style="padding:5px;background-color:$ErrorCol">$($ServiceApp.Status)</td>
"@
            }
            $emailbody += @"
                </tr>

"@
}

$emailbody += @"
                    </table>
                    <br />
"@


$ReportsList = [Microsoft.SharePoint.Administration.Health.SPHealthReportsList]::Local  
$items =  $ReportsList.Items | Where {$_['Severity'] -ne '4 - Success'}

$emailBody += @"
  <h2>SharePoint Health Analyser Reports</h2>
  <p style='font-size:8pt;'>$SPHAURL</p>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
    <tr bgcolor="#3366cc" style="color:#ffffff">
    <th style="padding:5px">Severity</th>
    <th style="padding:5px">Category</th>
    <th style="padding:5px">Title</th>
    <th style="padding:5px">Failing Servers</th>
    <th style="padding:5px">Failing Services</th>
    <th style="padding:5px">Modified</th>
    <th style="padding:5px">Expalanation</th>
   </tr>
"@

foreach ($item in $items) {
            if($Col -eq "#fff") {
                $Col = "#ddd"
            }
            else {
                $Col = "#fff"
            }
            $emailbody += @"
                <tr style='background-color:$Col'>
"@
            
            if ($item['Severity'] -eq "1 - Error") {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$ErrorCol">$($item['Severity'] -replace '. - ','')</td>

"@
            } elseif ($item['Severity'] -eq "4 - Success") {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$SuccessCol">$($item['Severity'] -replace '. - ','')</td>

"@
            } else {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$WarningCol">$($item['Severity'] -replace '. - ','')</td>

"@
            }
            $emailbody += @"
                <td style="padding:5px">$($item['Category'])</td>
                <td style="padding:5px">$($item.title)</td>
"@
            if ($item['Failing Servers']) {
                $emailbody += @"
                    <td style="padding:5px">$($item['Failing Servers'])</td>
"@
            } else {
                $emailbody += @"
                    <td style="padding:5px">n/a</td>
"@
            }
            if ($item['Failing Services']) {
                $emailbody += @"
                    <td style="padding:5px">$($item['Failing Services'])</td>
"@
            } else {
                $emailbody += @"
                    <td style="padding:5px">n/a</td>
"@
            }
            $emailbody += @"
                <td style="padding:5px">$($item['Modified'])</td>
                <td style="padding:5px">$($item['Explanation'])</td>
                </tr>
"@
}

$emailbody += @"
                    </table>
                    <br />
"@


$farm = Get-SPFarm
$cacheServices = $farm.Services | where {$_.Name -eq "AppFabricCachingService"}
$DCServers = $spservers | where {$_.Role -eq "DistributedCache"}

$scriptblock = {
          Invoke-Command {reg export HKLM\SOFTWARE\Microsoft\AppFabric\V1.0\Configuration "d:\temp\AppFabBackup.reg" /y } | out-null
          $output = Get-Content "d:\temp\AppFabBackup.reg" -Raw
          $output
    }
$DCStatus = Invoke-Command $scriptblock -ComputerName $DCServers[0].Address -SessionOption $option

[string]$regfile = $tempfolder + "AppFabBackup.reg"
$DCStatus | out-file "$regfile"
Invoke-Command {reg import "$regfile"} | Out-Null

Use-CacheCluster
$Cachehosts = Get-CacheHost

$CachehostConfigs = @()

foreach ($Cachehost in $Cachehosts) {
    $CachehostConfigs += Get-AFCacheHostConfiguration -ComputerName $Cachehost.HostName -CachePort $Cachehost.PortNo
}


$emailBody += @"
  <h2 style='font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>Distributed Cache Health</h2>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
 <tr bgcolor="#3366cc" style="color:#ffffff">
  <th style="padding:5px">Name</th>
  <th style="padding:5px">Server</th>
  <th style="padding:5px">CachePort</th>
  <th style="padding:5px">Version</th>
  <th style="padding:5px">Cache Size</th>
  <th style="padding:5px">Status</th>
 </tr>
"@

$Col = "#ddd"

foreach ($cacheService in $cacheServices) {
        if($Col -eq "#fff") {
            $Col = "#ddd"
        }
        else {
            $Col = "#fff"
        }
        $emailbody += @"
                  <tr style='background-color:$Col'>
                  <td style="padding:5px">$($cacheService.Name) SharePoint Service</td>
                  <td style="padding:5px">n/a</td>
                  <td style="padding:5px">n/a</td>
                  <td style="padding:5px">$($cacheService.Version)</td>
                  <td style="padding:5px">n/a</td>
"@
        if ($cacheService.Status -eq "Online") {
            $emailbody += @" 
                  <td style="padding:5px;background-color:$SuccessCol">$($cacheService.Status)</td>
"@ 
        } else {
            $emailbody += @" 
                  <td style="padding:5px;background-color:$ErrorCol">$($cacheService.Status)</td>
"@ 
        }
        $emailbody += @"
                  </tr>
"@    
}

foreach ($Cachehost in $Cachehosts) {
        if($Col -eq "#fff") {
            $Col = "#ddd"
        }
        else {
            $Col = "#fff"
        }
        $emailbody += @"
                  <tr style='background-color:$Col'>
                  <td style="padding:5px">$($Cachehost.ServiceName)</td>
                  <td style="padding:5px">$($Cachehost.HostName)</td>
                  <td style="padding:5px">$($Cachehost.PortNo)</td>
                  <td style="padding:5px">$($Cachehost.VersionInfo)</td>
"@
        foreach ($CachehostConfig in $CachehostConfigs) {
            if ($CachehostConfig.HostName -eq $Cachehost.HostName) {
            $emailbody += @"
                  <td style="padding:5px">$("{0:N2}" -f $($CachehostConfig.Size)) MB</td>
"@
            }
        }
        if ($Cachehost.Status -eq "Up") {
            $emailbody += @"
                  <td style="padding:5px;background-color:$SuccessCol">$($Cachehost.Status)</td>
                  </tr>
"@
        } else {
            $emailbody += @"
                  <td style="padding:5px;background-color:$ErrorCol">$($Cachehost.Status)</td>
                  </tr>
"@    
        }
}

$emailbody += @"
                    </table>
                    <br />
"@



$searchcomps = Get-SPEnterpriseSearchServiceApplication | Get-SPEnterpriseSearchStatus

$emailBody += @"
  <h2 style='font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>Search Topology Health</h2>
  <p style='font-size:8pt;'>$AdministrationURL</p>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
 <tr bgcolor="#3366cc" style="color:#ffffff">
  <th style="padding:5px">Server Name</th>
  <th style="padding:5px">Component</th>
  <th style="padding:5px">Health</th>
 </tr>
"@

$Col = "#ddd"


foreach ($searchcomp in $searchcomps) {
        if($Col -eq "#fff") {
            $Col = "#ddd"
        }
        else {
            $Col = "#fff"
        }
        $emailbody += @"
                                <tr style='background-color:$Col'>
"@
        if ($searchcomp.Details["Host"]) {
            $emailbody += @"
                                 <td style="padding:5px">$($searchcomp.Details["Host"])</td>
"@
        } else {
            $emailbody += @"
                                 <td style="padding:5px">n/a</td>
"@
        }
        $emailbody += @"
                                 <td style="padding:5px">$($searchcomp.Name)</td>
"@
        if ($searchcomp.State -eq "Active") {
            $emailbody += @"
                                 <td style="padding:5px;background-color:$SuccessCol">$($searchcomp.State)</td>
"@
        } else {
            $emailbody += @"
                                 <td style="padding:5px;background-color:$ErrorCol">$($searchcomp.State)</td>
"@
        }
        $emailbody += @"
                                </tr>
"@
    
    
}

$emailbody += @"
                    </table>
                    <br />
"@


$ServicesOnServers = @()

foreach ($spserver in $spservers) {
    $ServicesOnServer = get-spserver $spserver.Address | select ServiceInstances -ExpandProperty ServiceInstances | Sort-Object TypeName
    $Services =@()
    $count = 1
    foreach ($service in $ServicesOnServer) {
        $object = New-Object -TypeName PSObject
        
              if ($Service.TypeName -eq "SharePoint Server Search") {
                $object | Add-Member -Name 'TypeName' -MemberType Noteproperty -Value ($Service.TypeName + "_" + $count)
                $count++
              } else {
                $object | Add-Member -Name 'TypeName' -MemberType Noteproperty -Value $Service.TypeName
              }
              $object | Add-Member -Name 'Status' -MemberType Noteproperty -Value $Service.Status
              $object | Add-Member -Name 'ID' -MemberType Noteproperty -Value $Service.Id
        
        $Services += $Object
    }

    $ParentObject = New-Object PSObject -Property @{
                Server       = $spserver.Address
                Services     = $Services       
    }
    $ServicesOnServers += $ParentObject
}


$JsonConfig = "$tempfolder" + "ServicesOnServerConfig.json"

# the below will capture the current config as a baseline:
# $ServicesOnServers | ConvertTo-Json -depth 10 | out-file $JsonConfig

$ServicesOnServersConfig = Get-Content $JsonConfig | ConvertFrom-Json

$serviceslist = @()

foreach ($ServicesOnServer in $ServicesOnServers) {
    foreach ($Service in $ServicesOnServer.Services) {
        $Object = New-Object PSObject -Property @{
            TypeName = $Service.TypeName
        }
        $serviceslist += $Object
    }
}

$serviceslist = $serviceslist | Select-Object TypeName -Unique



$ServerServices = @()

foreach ($service in $serviceslist) {
    $object = New-Object -TypeName PSObject
    $object | Add-Member -Name 'Service' -MemberType Noteproperty -Value $service.TypeName
    foreach ($Server in $ServicesOnServersConfig) {
        foreach ($Servo in  $Server.Services) {
            if ($Servo.TypeName -eq $service.TypeName) {
                $object | Add-Member -Name $Server.Server -MemberType Noteproperty -Value $Servo.Status
            }
        }
    }
    $ServerServices += $object
}

$ServerServices = $ServerServices | Sort-Object Service


$emailBody += @"
  <h2 style='font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>Services on Server (Service Instances)</h2>
  <p style='font-size:8pt;'>$SPServicesOnServerURL</p>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
 <tr bgcolor="#3366cc" style="color:#ffffff">
  <th style="padding:5px">Service</th>
"@

foreach ($spserver in $spservers) {
$emailBody += @"
  <th style="padding:5px">$($spserver.Address)</th>
"@
}


$emailBody += @"
 </tr>
"@

$Col = "#ddd"

foreach ($ServerService in $ServerServices) {
if($Col -eq "#fff") {
            $Col = "#ddd"
        }
        else {
            $Col = "#fff"
        }
        $emailbody += @"
                                <tr style='background-color:$Col'>
"@
        $emailbody += @"
                                 <td style="padding:5px">$($ServerService.Service)</td>
"@
        foreach($spserver in $spservers) {
            $value = $ServicesOnServersConfig | where {$_.Server -eq $spserver.Address}
            $value = $value.Services | where {$_.TypeName -eq $ServerService.Service }
            if ($ServerService.($spserver.Address) -eq 1) {
                if ($value.Status -eq $ServerService.($spserver.Address)) {                
                    $emailbody += @"
                        <td style="padding:5px;background-color:$SuccessCol">Stopped</td>
"@
                } else {
                    $emailbody += @"
                        <td style="padding:5px;background-color:$ErrorCol">Stopped</td>
"@
                }
            } elseif ($ServerService.($spserver.Address) -eq 0) {
                if ($value.Status -eq $ServerService.($spserver.Address)) {                
                    $emailbody += @"
                        <td style="padding:5px;background-color:$SuccessCol">Started</td>
"@
                } else {
                    $emailbody += @"
                        <td style="padding:5px;background-color:$ErrorCol">Started</td>
"@
                }
            } else {
                $emailbody += @"
                    <td style="padding:5px">n/a</td>
"@
            }
        }


        $emailbody += @"
                                </tr>
"@
    
    
}

$emailbody += @"
                    </table>
                    <br />
"@






$emailBody += @"
  <h2>Crawl Log - Crawl History (Last 24 hours)</h2>
  <p style='font-size:8pt;'>$CrawlLogURL</p>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
 <tr bgcolor="#3366cc" style="color:#ffffff">
  <th style="padding:5px">Content Source</th>
  <th style="padding:5px">Started</th>
  <th style="padding:5px">Completed</th>
  <th style="padding:5px">Duration</th>
  <th style="padding:5px">Type</th>
  <th style="padding:5px">Successes</th>
  <th style="padding:5px">Warnings</th>
  <th style="padding:5px">Errors</th>
  <th style="padding:5px">Top Level Errors</th>
  <th style="padding:5px">Deletes</th>
  <th style="padding:5px">Not modified</th>
  <th style="padding:5px">Security updates</th>
  <th style="padding:5px">Security errors</th>
  <th style="padding:5px">Average Crawl Rate (dps)</th>
  <th style="padding:5px">Repository Latency (ms)</th>
 </tr>
"@

$Col = "#ddd"

foreach ($ContentSource in $ContentSources) {

    #Get Crawl History
    $CrawlLog = new-object Microsoft.Office.Server.Search.Administration.CrawlLog($SSA)
    $CrawlHistoryRows = $CrawlLog.GetCrawlHistory($ResultsCount, $ContentSource.Id)


    foreach ($CrawlHistoryRow in $CrawlHistoryRows) {
        if($Col -eq "#fff") {
            $Col = "#ddd"
        }
        else {
            $Col = "#fff"
        }
        if ([System.TimeZoneInfo]::ConvertTimeFromUtc($CrawlHistoryRow.CrawlStartTime, $TZ) -gt (Get-Date).AddHours(-24) ) {
            $CrawlDuration += $CrawlHistoryRow.CrawlDuration
            $CrawlRate += $CrawlHistoryRow.AverageCrawlRate
            $emailbody += @"
                                <tr style='background-color:$Col'>
                                 <td style="padding:5px">$($CrawlHistoryRow.ContentSourceName)</td>
                                 <td style="padding:5px">$([System.TimeZoneInfo]::ConvertTimeFromUtc($CrawlHistoryRow.CrawlStartTime, $TZ))</td>
"@          
            if ([System.TimeZoneInfo]::ConvertTimeFromUtc($CrawlHistoryRow.CrawlStartTime, $TZ) -eq $null ) {           
                $emailbody += @"
                                 <td style="padding:5px">In Progress</td>
"@  
            } else {
                $emailbody += @"
                                 <td style="padding:5px">$([System.TimeZoneInfo]::ConvertTimeFromUtc($CrawlHistoryRow.CrawlEndTime, $TZ))</td>
"@  
            }

            if (($CrawlHistoryRow.CrawlDuration.totalminutes -gt 15) -and  ($CrawlHistoryRow.CrawlDuration.totalminutes -lt 20)) {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$WarningCol">$("{0:hh\:mm\:ss}" -f ([TimeSpan] $CrawlHistoryRow.CrawlDuration))</td>

"@
            } elseif (($CrawlHistoryRow.CrawlDuration.totalminutes -gt 20) -or ($CrawlHistoryRow.CrawlDuration.totalminutes -eq 20)) {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$ErrorCol">$("{0:hh\:mm\:ss}" -f ([TimeSpan] $CrawlHistoryRow.CrawlDuration))</td>

"@
            } else {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$SuccessCol">$("{0:hh\:mm\:ss}" -f ([TimeSpan] $CrawlHistoryRow.CrawlDuration))</td>

"@
            }
            if ($CrawlHistoryRow.CrawlType -eq 2) {
                $emailBody += @"
                                 <td style="padding:5px">Incremental</td>
"@
            } else {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$WarningCol">Full</td>
"@
            }

            $emailbody += @"
                                 <td style="padding:5px">$($CrawlHistoryRow.Successes.ToString('N0'))</td>
                                 <td style="padding:5px">$($CrawlHistoryRow.Warnings.ToString('N0'))</td>
                                 <td style="padding:5px">$($CrawlHistoryRow.Errors.ToString('N0'))</td>
                                 <td style="padding:5px">$($CrawlHistoryRow.TopLevelErrors.ToString('N0'))</td>
                                 <td style="padding:5px">$($CrawlHistoryRow.Deletes.ToString('N0'))</td>
                                 <td style="padding:5px">$($CrawlHistoryRow.NotModified.ToString('N0'))</td>
                                 <td style="padding:5px">$($CrawlHistoryRow.SecurityUpdates.ToString('N0'))</td>
                                 <td style="padding:5px">$($CrawlHistoryRow.SecurityErrors.ToString('N0'))</td>
"@            
            
            if (($CrawlHistoryRow.AverageCrawlRate -gt 1) -and ($CrawlHistoryRow.AverageCrawlRate -lt 2) ) {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$WarningCol">$([math]::Round($CrawlHistoryRow.AverageCrawlRate,2))</td>
"@
            } elseif (($CrawlHistoryRow.AverageCrawlRate -lt 1) -or ($CrawlHistoryRow.AverageCrawlRate -eq 1)) {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$ErrorCol">$([math]::Round($CrawlHistoryRow.AverageCrawlRate,2))</td>
"@
            } else {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$SuccessCol">$([math]::Round($CrawlHistoryRow.AverageCrawlRate,2))</td>
"@
            }
            if (($CrawlHistoryRow.AverageRepositoryTime -gt 0) -and ($CrawlHistoryRow.AverageRepositoryTime -lt 1) ) {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$WarningCol">$($CrawlHistoryRow.AverageRepositoryTime)</td>
"@
            } elseif (($CrawlHistoryRow.AverageRepositoryTime -gt 1) -or ($CrawlHistoryRow.AverageRepositoryTime -eq 1)) {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$ErrorCol">$($CrawlHistoryRow.AverageRepositoryTime)</td>
"@
            } else {
                $emailbody += @"
                                 <td style="padding:5px;background-color:$SuccessCol">$($CrawlHistoryRow.AverageRepositoryTime)</td>
"@
            }
            $emailbody += @"
                                </tr>
"@
        }
    }


}




$CrawlDurationAverage = ($CrawlDuration | Measure-Object -Average -Property Ticks ).Average
$CrawlDurationAverage = [timespan]::FromTicks($CrawlDurationAverage)
$CTH = $CrawlDurationAverage.Hours
$CTM = $CrawlDurationAverage.Minutes
$CTS = $CrawlDurationAverage.Seconds
$CTMs = $CrawlDurationAverage.Milliseconds

$CrawlRateAverage = $CrawlRate | Measure-Object -Average

$SearchableItems = $CrawlLog.GetCrawledURLCount("",$false,-1,0,-1,[datetime]::minvalue,[datetime]::maxvalue)


$emailbody += @"
                    </table>
                    <p>Average Crawl Duration: $CTH hrs, $CTM min, $CTS sec, $CTMs ms<br />
                    Average Crawl Rate: $([math]::Round($CrawlRateAverage.Average,2)) documents per second<br />
                    Searchable Items: $($SearchableItems.ToString('N0'))
                    </p>
                    <br />
"@




$emailBody += @"
  <h2 style='font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>Crawl Log - Error Breakdown</h2>
  <p style='font-size:8pt;'>$CrawlErrorURL</p>
  <table border="0" cellpadding="3" style='font-size:10pt;font-family:"Segoe UI","Segoe",Tahoma,Helvetica,Arial,sans-serif;color:#555'>
 <tr bgcolor="#3366cc" style="color:#ffffff">
  <th style="padding:5px">Count</th>
  <th style="padding:5px">Error Message</th>
 </tr>
"@

$Col = "#ddd"
$CrawlErrorsTotal = @()

foreach ($ContentSource in $ContentSources) {
    $CrawErrors = $CrawlLog.GetCrawlErrors($ContentSource.Id,1)
    Foreach ($CrawError in $CrawErrors) { 
        $CrawlErrorsTotal += $CrawError.ErrorCount
        if($Col -eq "#fff") {
            $Col = "#ddd"
        }
        else {
            $Col = "#fff"
        }
        $emailbody += @"
                                <tr style='background-color:$Col'>
"@
        if ($CrawError.ErrorCount -gt 100) {
            $emailbody += @"
                                 <td style="padding:5px;background-color:$ErrorCol">$($CrawError.ErrorCount)</td>
"@
        } else {
            $emailbody += @"
                                 <td style="padding:5px;background-color:$WarningCol">$($CrawError.ErrorCount)</td>
"@
        }
        $emailbody += @"
                                 <td style="padding:5px">$($CrawError.ErrorMessage)</td>
                                </tr>
"@
    }
}

$CrawlErrorsTotal = $CrawlErrorsTotal | Measure-Object -Sum

$emailbody += @"
                    </table>
                    <p>Total Crawl Errors: $($CrawlErrorsTotal.Sum)</p>
"@





$RunTime = (Get-Date) - $startTime
$RTH = $RunTime.Hours
$RTM = $RunTime.Minutes
$RTS = $RunTime.Seconds
$RTMs = $RunTime.Milliseconds

$emailbody += @"
                    <br />
                    ____________________________________________________________________________
                    <p style="font-size:9pt">Script executed from: $myFQDN</p>
                    <p style="font-size:9pt">Total script execution time: "$RTH hrs $RTM min $RTS sec"</p>
                    </body>
                    </html>
"@

 
 
#Get outgoing Email Server
$EmailServer = (Get-SPWebApplication -IncludeCentralAdministration | Where { $_.IsAdministrationWebApplication } ) | %{$_.outboundmailserviceinstance.server} | Select Address
 
$From = "FromEmail@domain.com"
$To = "ToEmail@domain.com"
$Subject = "SharePoint Server Health Report - "+$CurrentDate
 
#Send Email
Send-MailMessage -smtpserver $EmailServer.Address -from $from -to $to -subject $subject -body $emailbody -BodyAsHtml