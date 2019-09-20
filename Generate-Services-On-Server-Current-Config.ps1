Add-PSSnapin Microsoft.SharePoint.Powershell -ErrorAction SilentlyContinue

$tempfolder = "d:\temp\"

$spservers = Get-SPServer | where {$_.Role -ne "Invalid"}

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

$ServicesOnServers | ConvertTo-Json -depth 10 | out-file $JsonConfig