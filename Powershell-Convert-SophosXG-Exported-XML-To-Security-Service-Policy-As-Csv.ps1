
#Build result folder in current folder.
$CurrentPS1File = $(Get-Item -Path "$PSCommandPath")
Set-Location "$($CurrentPS1File.PSParentPath)"
$OutputPath = $($CurrentPS1File.PSParentPath) + '\' + $($CurrentPS1File.BaseName) + '-CsvTableFile\' + $($CurrentPS1File.BaseName) + '-SecurityPolicyTable-' + $(Get-Date).ToString('yyyyMMdd-HHmmss') + '.csv'
$OutputPath = $OutputPath.replace("\\","\")
New-Item -ItemType Directory -Force -Path "$($($CurrentPS1File.PSParentPath) + '\' + $($CurrentPS1File.BaseName) + '-CsvTableFile\')"

#Get XG data from exported Entities.xml file.
[xml]$XMLObj = Get-Content -Path  .\Entities.xml -Encoding UTF8
$Interfaces = $XMLObj.Configuration.Interface
$SecurityPolicys = $XMLObj.Configuration.SecurityPolicy
$IPHosts = $XMLObj.Configuration.IPHost
$Services = $XMLObj.Configuration.Services

#Filting with mapped port.
$TargetTCPPort = "*"
$TargetServices = $Services | Where-Object{$_.ServiceDetails.ServiceDetail.DestinationPort -like $TargetTCPPort}

#Loop merge security policy with other data.
$Report = @()
for($SecurityPolicyIndex=0;$SecurityPolicyIndex -lt $SecurityPolicys.Count;$SecurityPolicyIndex++) {

    $Properties = @{}

    $MatchedService = $TargetServices | Where-Object{$_.Name -eq $SecurityPolicys[$SecurityPolicyIndex].Services.Service}

    if(($SecurityPolicys[$SecurityPolicyIndex].NonHTTPBasedPolicy.MappedPort -like $TargetTCPPort) -or (@($MatchedService).Count -gt 0)){

        $Name = $SecurityPolicys[$SecurityPolicyIndex].Name
        $Description = $SecurityPolicys[$SecurityPolicyIndex].Description
        $Zone = $SecurityPolicys[$SecurityPolicyIndex].NonHTTPBasedPolicy.SourceZones.Zone
        $HostedAddress = $SecurityPolicys[$SecurityPolicyIndex].NonHTTPBasedPolicy.HostedAddress
        $MappedPort = $SecurityPolicys[$SecurityPolicyIndex].NonHTTPBasedPolicy.MappedPort
        $ProtectedZone = $SecurityPolicys[$SecurityPolicyIndex].NonHTTPBasedPolicy.ProtectedZone
        $ProtectedServer = $SecurityPolicys[$SecurityPolicyIndex].NonHTTPBasedPolicy.ProtectedServer
        $ServiceName = $MatchedService.Name
        $DestinationPort = $MatchedService.ServiceDetails.ServiceDetail.DestinationPort
        $DestinationPort = $DestinationPort | ConvertTo-Json -Depth 100 | ConvertFrom-Json

        forEach($IPHost in $IPHosts){
            switch ($IPHost.HostType){
                'IP' {if($IPHost.Name -eq $ProtectedServer){$HostIP = $IPHost.IPAddress}}
                'IPRange' {}
                'IPList' {}
                'Network' {}
                'System Host' {
                    if($IPHost.Name -eq $HostedAddress){
                        $InterFaceName = $HostedAddress -replace '#([^#:]+)[:]?[^:]*','$1'
                        $PublicIP = ($Interfaces | ?{$_.Name -eq $InterFaceName}).IPAddress
                    }
                }
            }
        }

        $Properties = [ordered]@{
            Index = $SecurityPolicyIndex;
            Name = $Name;
            Description = $Description;
            Zone = $Zone;HostedAddress = $HostedAddress;
            PublicIP = $PublicIP;
            MappedPort = $MappedPort;
            ProtectedZone = $ProtectedZone;
            ProtectedServer = $ProtectedServer;
            HostIP = $HostIP;
            ServiceName = $ServiceName;
            DestinationPort = $DestinationPort
        }
        $Report += New-Object -TypeName PSObject -Property $Properties

    }
}
#Display security service policy table.
$Report | Format-Table -Property *
#Output security service policy table to csv file.
$Report | Export-Csv -Encoding UTF8 -Path $OutputPath
