<#
.TODO
Update Exchange Admin version https://www.alitajran.com/find-exchange-version-with-powershell/
Get virtual directories, then query DNS
.SYNOPSIS
	Document Exchange Environment ON prem.
	Gather server info
	This has been tested on Exchange 2013,2016,2019 with powershell 5 and higher
	This will mostly work with Exchange 2010
	File name:
		ExchangeOrganizationalDocumentation1.6.ps1
.COMPONENT
	RunAs Administrator
	Powershell Version 5 or higher
	Active directory module, the script will attempt to install and import it
	PsSession to exchange, will attempt to connect to Exchange Powershell virtual directory
.Notes
	Verifies the report path is accessible and will create the folders if needed
	This takes quite a while to run, patience grasshopper, There are 46 Reports here
.EXAMPLE
	Carlos Espinosa
.EXAMPLE
	3-22-2019
.EXAMPLE
	9-13-2019
		Added checks for Active Directory module
		added checks for Exchange Powershell
		Added check for Report folder access and creation
	10-2-2019
		Added Added reports for shared mailboxes
		Added CSV's to use for migration batches for shared mailboxes and mailboxes with SendAs permissions
	10-1-2020
		Updated to connect with powershell virtual directory
		Added additional field to mailbox details report, size in bytes
	10-13-2020
		Added netwroking details to Exchange server report
		Added Archive Mailbox stats report if archive exist
	10-26-2020
		Added Organizational configuration Report
		Added basic Dag Report
		Added MX record Report
		Added Autodiscover report
		Added DNS Resolution for virtual directory URL's
	1-11-2021
		Changed the way the script connects to Exchange to check for both PS virtual directory, if that fails try EMS
		Added CAS Array Report
	11-16-2021
		Added JournalRecipient, CircularLoggingEnabled, DatabaseGroup, Organization to Database report
		Updated virtual directory reports to look up public DNS records
	5-5-2022
		Added Public folder report
	5-11-2022
		Added $LocalDnsServer to attempt to gather the DNS servers that are configured on the server
	8-17-2022
		Added report to find all configured email domains on mailboxes
	9-6-2022
		Added reports for Retention tags and policies
#>

### Update the ReportPsessionath variable below before running the script, include drive and folder.. ####

$ReportPath = "C:\ATG Temp\Exchange Reports\Pre Change"
$ExchangePowerShellURL = "http://server.domain.com/powershell"

### Specify DNS servers to Query, at least one internal and one external server
# Attempt to get Configured DNS server from the NIC
$LocalDnsServer = Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object {!($_.InterfaceAlias -like "Loopback*")}
#$DNSServers = ("1.1.1.1","10.1.2.201") # this can be IP address or DNS name, this is used for autodiscover and MX records

$PublicDNSServer = "dns.google" # this is used to search public DNS records for the virtual directories
$DNSServers = ("$PublicDNSServer","$($LocalDnsServer.ServerAddresses[0])") # this is used for autodiscover and MX record lookup


### Reports that will be created   These are all CSV's #####
$ExchangeReportName = "Exchange Server Report.csv" ## Contains exchange servers, name, build number, exchange version, Exchange roles, creation date, Changed date, network addresses   ###
$OwaVirtDirReportName = "Virtual Directory OWA.csv" # Virtual directory Report
$EcpVirtDirReportName = "Virtual Directory ECP.csv"# Virtual directory Report
$ActSyncVirtDirReportName = "Virtual Directory ActiveSync.csv" # Virtual directory Report
$EwsVirtDirReportName = "Virtual Directory EWS.csv" # Virtual directory Report
$OabVirtDirReportName = "Virtual Directory OAB.csv" # Virtual directory Report
$MapiVirtDirReportName = "Virtual Directory Mapi.csv" # Virtual directory Report
$oAnyVirtDirReportName = "Virtual Directory OAnywhere.csv" # Virtual directory Report
$AutoDiscVirtDirReportName = "Virtual Directory Auto Discover.csv" # Virtual directory Report
$PsVirtDirReportName = "Virtual Directory PowerShell.csv" # Virtual directory Report
$MailBoxStatsReportName = "Mail Box Details.csv" #Mailbox stats user,Account enabled? ,Mailbox Alias,Is Resource,Is Shared,Primary SMTP, Mailbox Guid,DB Name, Mailbox Server, Mailbox Type,Mailbox Size
$SmtpAddressReportName = "Smtp Addresses Configured details.csv"  #list all smtp addresses for user mailboxes
$MailboxesWithForwarding = "Mailboxes With forwarding configured.csv" # list any forwarders configured on mailboxes
$MailboxSendAsAccess = "Mailboxes with Send-as Access.csv"  ### Report Where mailbox has send as Access
$MailboxSendOnBehalf = "Mailboxes with Send of Behalf.csv" ## Report on mailboxes with send on behalf
$MailboxFullAccess = "Mailbox Full Access Granted.csv"  ## mailboxes where full access has been granted
$MailboxSendAsAccessReportName = "Mailbox with SendAs permissions.csv" ### Report Where mailbox has send as Access
$MigrationBatchSendAsName = "Migration Batch SendAs.csv" # Csv File to create migration batch to office 365
$SharedMailBoxReportName = "Shared Mailbox Report.csv"  # list of shared mailboxes
$NestedGroupReportName = "Shared Mailbox Nested group Report.csv" # List of shared mailboxes with nested groups
$ShareMailBoxGroupMemberReportName = "Shared Mailbox Group membership Report.csv" # list of shared mailboxes with group membership
$MigrationBatchSharedReportName = "Migration Batch Shared mailboxes.csv"  # Csv file to create a migration batch to Office 365
$FullAccessReportName = "Mailbox Full Access Report.csv"
$FullAccessNestedGroupReportName = "Mailbox Full Access Nested group Report.csv"
$FullAccessGroupMemberReportName = "Mailbox Full Access Group membership Report.csv"
$MigrationBatchFullAccessReportName = "Migration Batch Full Access mailboxes.csv"
$MailboxSendOnBehalfReportName = "Mailbox SendOnBehalf Report.csv"
$NestedGroupSendOnBehalfReportName = "Mailbox Nested group SendOnBehalf Report.csv"
$MailboxGroupMemberSendOnBehalfReportName = "Mailbox Group membership SendOnBehalf Report.csv"
$MigrationBatchSendOnBehalfReportName = "Migration Batch SendOnBehalf Mailboxes.csv"
$OrganizationConfigReportName = "Oragnizational Config.CSV"
$DagReportName = "Database DagReport.csv"
$MXReportName = "DNS MX Records.csv"
$AutoDiscoverDnsReportName = "DNS AutoDiscover Records.csv"
$ExchangeHDWReportName = "Exchange Server Hardware.csv"
$AcceptedDomainReportName = "Exchange Accepted Domains.csv"
$ClientAccessReportName = "Client Access.csv"
$ReveiveConnectorReportName = "Connectors Receive Connectors.csv"
$CasArrayReportName = "Cas Arrays.csv"
$SendConnectorReportName = "Connectors Send Connectors.csv"
$ExchangeCertReportName = "Exchange Certificates.csv"
$ExchangeDataBaseReportName = "Exchange Database Config.csv"
$ExchangeAdminReportName = "Exchange Admin Accounts.csv"
$PublicFolderStatsReportName = "PublicFolder Details.csv" 
$RetentionTagReportName = "Retention Tags.csv" 
$RetentionPolicyReportName = "Retention Policies.csv" 
########################################################

## Reports - TXT format
$MailboxCountName = "MailBox Count and Types.txt" # List of mail box types ,(User,Room,Shared...) with count
$ClientAccessServerTxt = "Client Access Servers.txt"  ## show servers that host autodiscover and Outlook anywhere
$EmailAddressPolicy = "Email Address Policy.txt"  ##
$ReceiveConnectorTXTRpt = "Connector Recieve Connectors.txt"
$TransportConfiguration = "Transport Configuration.txt"
$OWAPolicies ="OWA Policies.txt"
$MobileDevicePolicy = "Active Sync Mobile device policy.txt"
$TransportRules = "Transport rules.txt"


## nothing to change below ####
### Do not change the variable below as it is needed to prevent truncation of data in the output

$FormatEnumerationLimit=-1
# Verify the report path can be accessed and that the folder is present
if (!(Test-Path $ReportPath)) {
    Try {
        New-Item -ItemType Directory -Path $ReportPath -ErrorAction Stop
    }
    Catch {
        Write-Warning "Unable to access report path, verify the drive exists and that the account running his has permissions to create a folder"
        Return
    }
}
#check for Exchange PsSession
$PsSession = $null
$PsSession = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}
if (!($PsSession.Count -ge 1)) {
	Write-Warning "Exchange PsSession Not present attempting to connect to Exchange Management Shell"
	Try {
	# try to connect to Exchange managment shell
		$CallEMS = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell "
		Invoke-Expression $CallEMS
	}
	Catch {
		Write-host "Unable to connect to Exchange Managenment Shell.  Is Exchange installed on this computer?" -ForegroundColor Yellow
		Write-Warning "Connection to Exchange management shell failled, attempting to connect to Powershell virtual directory"
		Try {
		$UserCredential = Get-Credential
# Try to Connect to Exchange Powershell virtual Directory
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangePowerShellURL -Authentication Kerberos -Credential $UserCredential -ErrorAction SilentlyContinue
		Import-PSSession $Session -DisableNameChecking
		}
		Catch {
			Write-host "Unable to connect to Powershell virtual Directory at $ExchangePowerShellURL is it spelled correctly?" -foregroundcolor yellow
			Return
		}
	}
}
else {
	Write-host "Exchange Powershell is available."
}

# Check for AD Powershell
If (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Warning "ACtive Directory Module is not avalable, Attempting to install"
    Install-WindowsFeature RSAT-AD-PowerShell
}
Else {
    Import-Module ActiveDirectory
}

Function MX-Results {
	$MXReport = [PsCustomObject] [Ordered] @{
	"Accepted Domain" = "$($AcceptedDomain.DomainName)"
	"Domain DNS Exists" = $DomainRecordExists
	"MX Record Exists" = $MXExists
	"MX Name" = "$($MXRecord.Name)"
	"Record Type" = "$($MXRecord.QueryType)"
	"Txt" = "$($MXRecord.Strings)"
	"Exchange Name" = "$($MXRecord.NameExchange)"
	"TTL" = "$($MXRecord.TTL)"
	"Responding DNS Server" = $RespondingDNSServer
	}
	$MXReport | Export-Csv -NoTypeInformation -Append "$ReportPath\$MXReportName"
}

Function AutoDiscover-Report {
	$DNSReport = [PsCustomObject] [Ordered] @{
	"Accepted Domain" = "$($AcceptedDomain.DomainName)"
	"Domain DNS Record Exists" = $DomainRecordExists
	"Autodiscover name" = "$($DNSRecord.Name)"
	"Autodiscover Record Exists" = $AutoDiscoverRecordExits
	"Autodiscover Type" = "$($DNSRecord.Type)"
	"Autodiscover TTL" = "$($DNSRecord.TTL)"
	"Autodiscover namehost" = "$($DNSRecord.NameHost)"
	"IP Addresses" = $IpAddress
	"Responding DNS Server" = $RespondingDNSServer
	}
	$DNSReport | Export-Csv -NoTypeInformation -Append "$ReportPath\$AutoDiscoverDnsReportName"
}

Function ExchangeHDW-Report {
	$ExchangeHDWReport += [PsCustomObject] [Ordered] @{
	"Computer Name" = "$($Computer.ToUpper())"
	"IP Address" = "$($network.IpAddress -join ",")"
	"Subnet Mask" = "$($Network.IpSubnet -join ",")"
	"Default GateWay" = "$($Network.DefaultIPGateway -join ",")"
	"Primary DNS" = "$($PrimaryDNSServer)"
	"Secondary DNS" = "$($SecondaryDNSServer)"
	"Tertiary DNS" = "$($TertiaryDNSServer)"
	"MacAddress" = "$($network.MACAddress)"
	"IsDHCPEnabled" = "$($network.DHCPEnabled)"
	"Network Name" = "$($Network.Description)"
	"OS Version" = "$($OsInfo.Caption)"
	"OsArchitecture" = "$($OsInfo.OsArchitecture)"
	"OS Install Date" = "$($OSInstallDate)"
	"OsDirectory" = "$($OsInfo.windowsdirectory)"
	"Harware Roles" = "$($hardware.roles -join ",")"
	"LogicalProcessors" = "$($Hardware.NumberOfLogicalProcessors)"
	"Processors" = "$($Hardware.NumberOfProcessors)"
	"RAM" = "$($TotalMemory)"
	"Drive Letter" = "$($Disk.DeviceID)"
    "Drive Capacity (GB)" = "$($DriveCapacity)"
    "Drive FreeSpace (GB)" = "$($FreeSpace)"
    "Drive UsedSpace (GB)" = "$($UsedSpace)"
    "Drive Free (%)" = "$($FreeSpacePercent)%"
	"Accessible" = "$($Accessible)"
	"Dns Record" = "$($DnsIP)"
	}
	$ExchangeHDWReport | Export-Csv -Append -NoTypeInformation "$ReportPath\$ExchangeHDWReportName"
}
### Do not change the variable below as it is needed to prevent truncation of data in the output

$FormatEnumerationLimit=-1

### Do not change the variables below ###
$Mailboxes = get-mailbox -ResultSize Unlimited
$Servers = Get-ExchangeServer
$OrganizationConfig = Get-OrganizationConfig
$AcceptedDomains = Get-AcceptedDomain
$ClientAccessServers = Try{Get-ClientAccessService} Catch{Get-ClientAccessServer}
$CASArrays = Get-ClientAccessArray
$ExchangeCerts = Get-ExchangeCertificate
$Dags = Get-DatabaseAvailabilityGroup
$SendConnectors = Get-SendConnector
$ExchangeAdmins = Get-AdGroupMember -Identity "Organization Management"
$PublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited -ErrorAction silentlycontinue
$RetentionTags = Get-RetentionPolicyTag 
$RetentionPolicies = Get-RetentionPolicy 

########################################################

## Get exchange servers, name, build number, exchange version, Exchange roles, creation date, Changed date, network addresses   ###
## this still needs error checking
$ExchangeServerReport = @()
    foreach ($ExchangeServer in $servers) {
		$NicInfo = Get-NetworkConnectionInfo -Identity $ExchangeServer.Name
         $ExchangeServerReport += [PsCustomObject] [Ordered] @{
            "Name" = "$($ExchangeServer.Name)"
            "Domain" = "$($ExchangeServer.Domain)"
			"AD Site" = "$($ExchangeServer.site)"
			"Edition" = "$($ExchangeServer.Edition)"
			"FQDN" ="$($ExchangeServer.fqdn)"
			"Admin Display version" = "$($ExchangeServer.AdminDisplayVersion)"
            "Exchange version" = "$($ExchangeServer.ExchangeVersion)"
			"Server Role" = "$($ExchangeServer.ServerRole)"
			"When Created" = "$($ExchangeServer.WhenCreated)"
			"When Changed" = "$($ExchangeServer.WhenChanged)"
            "IpAddress" = "$($NicInfo.IPAddresses.ipaddressToString -join(","))"
			"DNSSettings" = "$($NicInfo.DNSServers -join(","))"
        }
    }
$ExchangeServerReport | Export-Csv -NoTypeInformation "$ReportPath\$ExchangeReportName"

## Exchange Hardware Report
$ExchangeHDWReport = @()
ForEach ($ExchangeServer in $Servers) {
	$Computer = $ExchangeServer.Name
	$Networks = $null
	$Network = $null
	$OSInstallDate = $null
	$PrimaryDNSServer= $null
	$SecondaryDNSServer = $null
	$TertiaryDNSServer = $null
	$IsDHCPEnabled = $null
	$OsInfo = $null
	$Hardware = $null
	$TotalMemory = $null
	$DnsIP = $null
	$DriveCapacity = $null
	$FreeSpace = $null
	$UsedSpace = $null
	$FreeSpacePercent = $null
	$FreeSpacePercent = $null
	$Disk = $null
	$Disks = $null

	if(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {
	$Accessible = "Server Accessed"
	try {
		$Networks = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -ComputerName $Computer -ErrorAction SilentlyContinue
        $disks = Get-WmiObject win32_logicaldisk -ComputerName $computer -Filter "Drivetype=3" -ErrorAction SilentlyContinue
		$OsInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer -ErrorAction SilentlyContinue
		$Hardware = Get-Wmiobject Win32_computerSystem -Computername $Computer -ErrorAction SilentlyContinue
		$totalMemory = [math]::round($Hardware.TotalPhysicalMemory/1024/1024/1024, 2)
		$OSInstallDate = "$($OsInfo.ConvertToDateTime($OsInfo.InstallDate))"
	}
	catch {
		Write-Verbose "Failed to Query $Computer. Error details: $_"
        $Accessible = "WMI Failed"
		($DnsIP = $(Resolve-DnsName -Name $Computer -ErrorAction SilentlyContinue).IPAddress -join ",")
		ExchangeHDW-Report
	}
		foreach($Network in $Networks) {
			If(!($Network.DNSServerSearchOrder)) {
				$PrimaryDNSServer = "Notset"
				$SecondaryDNSServer = "Notset"
			}
			elseif($Network.DNSServerSearchOrder.count -eq 1) {
				$PrimaryDNSServer = $Network.DNSServerSearchOrder[0]
				$SecondaryDNSServer = "Notset"
			}
			elseif ($Network.DNSServerSearchOrder.count -eq 2) {
				$PrimaryDNSServer = $Network.DNSServerSearchOrder[0]
				$SecondaryDNSServer = $Network.DNSServerSearchOrder[1]
			}
			elseif ($Network.DNSServerSearchOrder.count -eq 3) {
				$PrimaryDNSServer = $Network.DNSServerSearchOrder[0]
				$SecondaryDNSServer = $Network.DNSServerSearchOrder[1]
				$TertiaryDNSServer = $Network.DNSServerSearchOrder[2]
			}
			ForEach ($disk in $disks) {
				$DriveCapacity =[math]::round($disk.Size / 1gb,2)
				$FreeSpace = [math]::round($disk.FreeSpace / 1gb,2)
				$UsedSpace = ($DriveCapacity - $FreeSpace)
				$FreeSpacePercent = ($FreeSpace/$DriveCapacity)*100
				$FreeSpacePercent = [math]::round($FreeSpacePercent)
				($DnsIP = $(Resolve-DnsName -Name $Computer -Type A -ErrorAction SilentlyContinue).IPAddress -join ",")
				ExchangeHDW-Report
			}
		}
	}
	else {
		Write-Verbose "$Computer not reachable"
		$Accessible = "Not Accessible"
		IF ($DnsIP = $(Resolve-DnsName -Name $Computer -Type A -ErrorAction SilentlyContinue).IPAddress -join ",") {
		}
		Else {
			$DnsIP = "No DNS Record Found"
		}
		ExchangeHDW-Report
	}
}

## Get Receive connetors
$ReceiveConnectorReport = @()
    foreach ($ExchangeServer in $servers) {
		$ReceiveConnectors = Get-ReceiveConnector -Server $ExchangeServer.Name
		Foreach ($ReceiveConnector in $ReceiveConnectors) {
			$ReceiveConnectorReport += [PsCustomObject] [Ordered] @{
				"Name" = "$($ReceiveConnector.Name)"
				"Server" = "$($ReceiveConnector.FQDN)"
				"Enabled" = "$($ReceiveConnector.Enabled)"
				"ProtocolLoggingLevel" = "$($ReceiveConnector.ProtocolLoggingLevel)"
				"Max message Size" = "$($ReceiveConnector.MaxMessageSize)"
				"Bindings" ="$($ReceiveConnector.bindings -join(","))"
				"Auth Method" = "$($ReceiveConnector.AuthMechanism -join(", "))"
				"Remote IP" = "$($ReceiveConnector.RemoteIPRanges -join(", "))"
				"When Created" = "$($ReceiveConnector.WhenCreated)"
				"When Changed" = "$($ReceiveConnector.WhenChanged)"
			}
		}
    }
$ReceiveConnectorReport | Export-Csv -NoTypeInformation "$ReportPath\$ReveiveConnectorReportName"

### Organization Config Report
$OrganizationConfigReport = @()
$OrganizationConfigReport += [PsCustomObject] [Ordered] @{
	"Name" = "$($OrganizationConfig.Name)"
	"OrganizationId" = "$($OrganizationConfig.OrganizationId)"
	"HybridConfigurationStatus" = "$($OrganizationConfig.HybridConfigurationStatus)"
	"IsMixedMode" = "$($OrganizationConfig.IsMixedMode)"
	"MapiHttpEnabled" = "$($OrganizationConfig.MapiHttpEnabled)"
	"AdfsAuthenticationConfiguration" = "$($OrganizationConfig.AdfsAuthenticationConfiguration)"
	"AdfsIssuer" = "$($OrganizationConfig.AdfsIssuer)"
	"IntuneManagedStatus" ="$($OrganizationConfig.IntuneManagedStatus)"
	"AzurePremiumSubscriptionStatus" = "$($OrganizationConfig.AzurePremiumSubscriptionStatus)"
	"RealTimeLogServiceEnabled" = "$($OrganizationConfig.RealTimeLogServiceEnabled)"
	"WhenCreated" = "$($OrganizationConfig.WhenCreated)"
	"WhenChanged" = "$($OrganizationConfig.WhenChanged)"
	}
$OrganizationConfigReport | Export-Csv -NoTypeInformation "$ReportPath\$OrganizationConfigReportName"

# Get MX records for all accepted domains
$MXReport = @()
ForEach ($DNSServer in $DNSServers) {
	foreach ($AcceptedDomain in $AcceptedDomains) {
	$DomainRecordExists = $null
	$MXExists = $null
	$DNS = $null
	$MXRecord = $null
	$RespondingDNSServer = $null
	# DNS Check
		$DNS = Resolve-DnsName -Name $AcceptedDomain.DomainName -Server $DNSServer -Type SOA -DnsOnly -TcpOnly -NoHostsFile -ErrorAction SilentlyContinue
		If ($DNS) {
			$DomainRecordExists = "$True"
            $MXSearch = Resolve-DnsName -Name $AcceptedDomain.DomainName -Server $DNSServer -Type "ALL" -DnsOnly -TcpOnly -NoHostsFile -ErrorAction SilentlyContinue
            $RecordCount = 0
			foreach ($MXRecord in $MXSearch){
				If ($MXRecord.QueryType -eq "TXT" -or $MXRecord.QueryType -eq "MX") {
					$MXExists = "True"
					$RespondingDNSServer = $($DNS[0].PrimaryServer)
					MX-Results
				}
                Else {
                    $RecordCount++
                    If (!($RecordCount -gt 1)) {
						$MXRecord = $null
						$MXExists = "False"
						$RespondingDNSServer = $DNSServer
						MX-Results
                    }
                }
			}
		}
		Else {
			$DomainRecordExists = "False"
			$RespondingDNSServer = $DNSServer
			MX-Results
		}
	}
}
# Get Autodiscover records
$DnsReport = @()
foreach ($DNSServer in $DNSServers) {
	foreach ($AcceptedDomain in $AcceptedDomains) {
    $DomainRecordExists = $null
    $AutoDiscoverRecordExits = $null
    $DNS = $null
    $DNSRecord = $null
    $AutoDiscover  = $null
	$RespondingDNSServer = $null
# Domain DNS Check
		$DNS = Resolve-DnsName -Name $AcceptedDomain.DomainName -Server $DNSServer -Type SOA -DnsOnly -TcpOnly -NoHostsFile -ErrorAction SilentlyContinue
		If ($DNS) {
			$DomainRecordExists = "True"
# search for autodiscover records
			$AutoDiscover = Resolve-DnsName -Name "Autodiscover.$($AcceptedDomain.DomainName)" -Server $DNSServer -Type all -DnsOnly -TcpOnly  -ErrorAction SilentlyContinue
            $IpAddress = $null
			IF ($AutoDiscover) {
				$AutoDiscoverRecordExits = "True"
				foreach ($DNSRecord in $AutoDiscover) {
				$RespondingDNSServer = $($DNS[0].PrimaryServer)
						If ($DNSRecord.QueryType -eq "CNAME") {
							Write-host "CNAME Found" -ForegroundColor Yellow
							$IpAddress = (Resolve-DnsName -Name $($DNSRecord.NameHost) -Type A).ip4address -join ", "
							AutoDiscover-Report
						}
						else {
							$RespondingDNSServer = $($DNS[0].PrimaryServer)
							$IpAddress = $($DNSRecord.IPAddress)
							AutoDiscover-Report
						}
				}
			}
			Else {
				Write-Host "Autodiscover record for $($AcceptedDomain.DomainName) not found" -ForegroundColor red
				$AutoDiscoverRecordExits = "False"
				$RespondingDNSServer = $DNSServer
				AutoDiscover-Report
			}
		}
		Else{
			Write-host "Dns Record for $($AcceptedDomain.DomainName) not found on $($DnsServer)" -ForegroundColor Yellow
			$DomainRecordExists = "False"
            $IpAddress = $null
			$RespondingDNSServer = $DNSServer
			AutoDiscover-Report
		}
	}
}

# Get Accepted Domains
$AcceptedDomainReport = @()
foreach ($AcceptedDomain in $AcceptedDomains) {
	$AcceptedDomainReport += [PsCustomObject] [Ordered] @{
		"Name" = "$($AcceptedDomain.Name)"
		"Domain name" = "$($AcceptedDomain.DomainName)"
		"Domain Type" = "$($AcceptedDomain.DomainType)"
		"Default" = "$($AcceptedDomain.Default)"
		"is Coexistence Domain" = "$($AcceptedDomain.IsCoexistenceDomain)"
		"Initial Domain" ="$($AcceptedDomain.InitialDomain)"
		"When Created" = "$($AcceptedDomain.WhenCreated)"
		"Distinguished Name" = "$($AcceptedDomain.DistinguishedName)"
		"Federated Org Link" = "$($AcceptedDomain.FederatedOrganizationLink)"
	}
}

$AcceptedDomainReport | Export-Csv -NoTypeInformation "$ReportPath\$AcceptedDomainReportName"

#  Client access servers report
$ClientAccessReport = @()
foreach ($ClientAccessServer in $ClientAccessServers) {
	$ClientAccessReport += [PsCustomObject] [Ordered] @{
		"Name" = "$($ClientAccessServer.Name)"
		"FQDN" = "$($ClientAccessServer.Fqdn)"
		"ClientaccessArray" = "$($ClientAccessServer.ClientAccessArray)"
		"Outlook anywhere enabled" = "$($ClientAccessServer.OutlookAnywhereEnabled)"
		"Autodiscover Internal Uri" = "$($ClientAccessServer.AutoDiscoverServiceInternalUri)"
		"AutoDiscover CN" ="$($ClientAccessServer.AutoDiscoverServiceCN)"
		"AutoDiscover SiteScope" = "$($ClientAccessServer.AutoDiscoverSiteScope)"
		"Distinguished Name" = "$($ClientAccessServer.DistinguishedName)"
	}
}

$ClientAccessReport | Export-Csv -NoTypeInformation "$ReportPath\$ClientAccessReportName"

#CAS Array Report
$CasArrayReport = @()
foreach ($CasArray in $CASArrays) {
	$CasArray.Members.name
		$CasArrayReport += [PsCustomObject] [Ordered] @{
		"Name" = "$($CasArray.Name)"
		"FQDN" = "$($CasArray.Fqdn)"
		"Members" = "$($CasArray.Members.Name -join (", "))"
		"AD Site" = "$($CasArray.SiteName)"
	}
}
$CasArrayReport | Export-CSV -NoTypeInformation "$ReportPath\$CasArrayReportName"

# Get-DatabaseAvailabilityGroup(s), DagName,IPAddress,ReplicationPort,loadbalancedEnabled,DagNetworkName,DagServers,DagServerIpAddresses,DagServerState,DagWitnessServer,DagWitnessDir
$DagReport = @()

	if (!($Dags)) {
		Write-Host "No database Avilability groups exist" -ForegroundColor Yellow
	}
	else {
		ForEach ($Dag in $Dags) {
			$DagNetwork = Get-DatabaseAvailabilityGroupNetwork -Identity $Dag.Name
			$DagReport += [PsCustomObject] [Ordered] @{
				"DagName" = "$($dag.name)"
				"IPAddress" = "$($Dag.DatabaseAvailabilityGroupIpv4Addresses -join(", "))"
				"ReplicationPort" = "$($Dag.ReplicationPort)"
				"LoadBalaned" = "$($Dag.MailboxLoadBalanceEnabled)"
				"DagNetworkName" = "$($DagNetwork.Name)"
				"DagServers" = "$($DagNetwork.Interfaces.NodeName -Join (", "))"
				"DagServerIpAddresses" =  "$($DagNetWork.Interfaces.IPAddress.IPAddressToString -join (", "))"
				"DagServerState" = "$($DagNetwork.Interfaces.State -join (", "))"
				"DagWitnessServer" = "$($Dag.WitnessServer.Fqdn)"
				"DagWitnessDir" = "$($Dag.WitnessDirectory.PathName)"
			}
		}
	$DagReport | Export-CSV -NoTypeInformation "$ReportPath\$DagReportName"
	}

## Get send connectors create report

$SendConnectorReport = @()

Foreach ($SendConnector in $SendConnectors) {
	$SendConnectorReport += [PsCustomObject] [Ordered] @{
		"Name" = "$($SendConnector.Name)"
		"ConnectorType" = "$($SendConnector.ConnectorType)"
		"Enabled" = "$($SendConnector.Enabled)"
		"FQDN" = "$($SendConnector.FQDN)"
		"Port" = "$($SendConnector.Port)"
		"addressSpaces" ="$($SendConnector.addressSpaces)"
		"SmartHosts" = "$($SendConnector.SmartHosts -join(", "))"
		"SourceTransportServers" = "$($SendConnector.SourceTransportServers -join(", "))"
		"TlsAuthLevel" = "$($SendConnector.TlsAuthLevel)"
		"TlsDomain " = "$($SendConnector.TlsDomain)"
		"ProtocolLoggingLevel" = "$($SendConnector.ProtocolLoggingLevel)"
		"MaxMessageSize" = "$($SendConnector.MaxMessageSize)"
	}
}

$SendConnectorReport | Export-Csv -NoTypeInformation "$ReportPath\$sendConnectorReportName"

## Get Exchange certificates

$ExchangeCertReport = @()

Foreach ($ExchangeCert in $ExchangeCerts) {
	$ExchangeCertReport += [PsCustomObject] [Ordered] @{
		"CertificateDomains " = "$($ExchangeCert.CertificateDomains -join(", "))"
		"HasPrivateKey" = "$($ExchangeCert.HasPrivateKey)"
		"IsSelfSigned" = "$($ExchangeCert.IsSelfSigned)"
		"Issuer" = "$($ExchangeCert.Issuer)"
		"NotAfter" = "$($ExchangeCert.NotAfter)"
		"RootCAType" ="$($ExchangeCert.RootCAType)"
		"SerialNumber" = "$($ExchangeCert.SerialNumber)"
		"Services" = "$($ExchangeCert.Services)"
		"Status" = "$($ExchangeCert.Status)"
		"Subject" = "$($ExchangeCert.Subject)"
		"Thumbprint" = "$($ExchangeCert.Thumbprint)"
	}
}

$ExchangeCertReport | Export-Csv -NoTypeInformation "$ReportPath\$ExchangeCertReportName"

## Get Databases
$ExchangeDataBaseReport = @()
    foreach ($ExchangeServer in $servers) {
		$ExchangeDataBases = Get-MailboxDatabase -Server $ExchangeServer.Name
		Foreach ($ExchangeDataBase in $ExchangeDataBases) {
			$ExchangeDataBaseReport += [PsCustomObject] [Ordered] @{
				"Name" = "$($ExchangeDataBase.Name)"
				"ServerName" = "$($ExchangeDataBase.ServerName)"
				"MasterServerOrAvailabilityGroup" = "$($ExchangeDataBase.MasterServerOrAvailabilityGroup)"
				"OfflineAddressBook" = "$($ExchangeDataBase.OfflineAddressBook)"
				"EdbFilePath" = "$($ExchangeDataBase.EdbFilePath)"
				"LogFolderPath" ="$($ExchangeDataBase.LogFolderPath)"
				"ProhibitSendReceiveQuota" = "$($ExchangeDataBase.ProhibitSendReceiveQuota)"
				"ProhibitSendQuota" = "$($ExchangeDataBase.ProhibitSendQuota)"
				"DatabaseCopies" = "$($ExchangeDataBase.DatabaseCopies -join (", "))"
				"DatabaseGroup" = "$($ExchangeDataBase.DatabaseGroup)"
				"Organization" = "$($ExchangeDataBase.Organization)"
				"ActivationPreference" = "$($ExchangeDataBase.ActivationPreference -join (", "))"
				"ReplayLagTimes" = "$($ExchangeDataBase.ReplayLagTimes -join (", "))"
				"Servers" = "$($ExchangeDataBase.Servers -join (", "))"
				"JournalRecipient" = "$($ExchangeDataBase.JournalRecipient)"
				"CircularLoggingEnabled" = "$($ExchangeDataBase.CircularLoggingEnabled)"
			}
		}
    }
$ExchangeDataBaseReport | Export-Csv -NoTypeInformation "$ReportPath\$ExchangeDataBaseReportName"

# Get Exchange Admins
$ExchangeAdminReport = @()
foreach ($ExchangeAdmin in $ExchangeAdmins) {
	If ($exchangeadmin.objectClass -eq "user"){
		$Account = Get-ADUser -Identity $($exchangeadmin.SamAccountName) -Properties *
		$GroupMembership = @()
		foreach ($group in $Account.MemberOf) {
			$GroupMembership += $(Get-ADGroup -Identity $group).Name
		}
	}
	elseIf ($exchangeadmin.objectClass -eq "group"){
		$Account = Get-ADGroup -Identity $($exchangeadmin.SamAccountName) -Properties *
	}
	$ExchangeAdminReport += [PsCustomObject] [Ordered] @{
		"DisplayName" = "$($Account.DisplayName)"
		"UserPrincipalName" = "$($Account.UserPrincipalName)"
		"EmailAddress" = "$($Account.EmailAddress)"
		"SamAccountName" = "$($Account.SamAccountName)"
		"CanonicalName" ="$($Account.CanonicalName)"
		"proxyAddresses" = "$($Account.proxyAddresses -join (", "))"
		"Enabled" ="$($Account.Enabled)"
		"PasswordNeverExpires" = "$($Account.PasswordNeverExpires)"
		"PasswordLastSet" = "$($Account.PasswordLastSet)"
		"ObjectClass" = "$($Account.ObjectClass)"
		"whenCreated" = "$($Account.whenCreated)"
		"MemberOf" = "$($GroupMembership -join (", "))"
	}
}

$ExchangeAdminReport | Export-Csv -NoTypeInformation "$ReportPath\$ExchangeAdminReportName"
## Retention Tag Report

$RetentionTagReport = @()
foreach ($RetentionTag in $RetentionTags) {
         $RetentionTagReport += [PsCustomObject] [Ordered] @{
            "Name" = "$($RetentionTag.Name)"
			"RetentionEnabled" = "$($RetentionTag.RetentionEnabled)"
			"RetentionAction" = "$($RetentionTag.RetentionAction)"
			"AgeLimitForRetention" = "$($RetentionTag.AgeLimitForRetention)"
			"MoveToDestinationFolder" = "$($RetentionTag.MoveToDestinationFolder)"
			"TriggerForRetention" = "$($RetentionTag.TriggerForRetention)"
			"Type" = "$($RetentionTag.Type)"
			"Description" = "$($RetentionTag.Description)"
			"IsDefaultAutoGroupPolicyTag" = "$($RetentionTag.IsDefaultAutoGroupPolicyTag)"
			"SystemTag" = "$($RetentionTag.SystemTag)"
			"MessageClassDisplayName" = "$($RetentionTag.MessageClassDisplayName)"
			"MessageClass" = "$($RetentionTag.MessageClass)"
         }
	}
$RetentionTagReport | Export-Csv -NoTypeInformation "$ReportPath\$RetentionTagReportName"

#Retention Policy Report

$RetentionPolicyReport = @()
foreach ($Policy in $RetentionPolicies) {
         $RetentionPolicyReport += [PsCustomObject] [Ordered] @{
            "Name" = "$($Policy.Name)"
			"IsDefault" = "$($Policy.IsDefault)"
			"RetentionPolicyTagLinks" = "$($Policy.RetentionPolicyTagLinks -join ", ")"
         }
	}

$RetentionPolicyReport | Export-Csv -NoTypeInformation "$ReportPath\$RetentionPolicyReportName"

## Public Folder report
$PublicFolderStatsReport = @()

foreach ($PublicFolder in $PublicFolders) {
    foreach ($PublicFolderStat in (Get-PublicFolderStatistics -identity $PublicFolder.Identity)) {
		$PublicFolderStatsReport += [PsCustomObject] [Ordered] @{
		"Identity" = "$($PublicFolder.Identity)"
		"Name" = "$($PublicFolder.Name)"
		"ContentMailboxName" = "$($PublicFolder.ContentMailboxName)"
		"FolderSize bytes" = "$($PublicFolder.FolderSize)"
		"MailboOwnerId" = "$($PublicFolder.MailboOwnerId)"
		"IsValid" = "$($PublicFolder.IsValid)"
		"ObjectState" = "$($PublicFolder.ObjectState)"
		"CreationTime" = "$($PublicFolderStat.CreationTime)"
		"LastModificationTime" = "$($PublicFolderStat.LastModificationTime)"
		"ItemCount" = "$($PublicFolderStat.ItemCount)"
		"DeletedItemCount" = "$($PublicFolderStat.DeletedItemCount)"
		"AssociatedItemCount" = "$($PublicFolderStat.AssociatedItemCount)"
		"TotalItemSize" = "$($PublicFolderStat.TotalItemSize)"
		}
	}
}
$PublicFolderStatsReport | Export-Csv -NoTypeInformation "$ReportPath\$PublicFolderStatsReportName"


### The following will create a report that includes user,Account enabled? ,Mailbox Alias,Is Resource,Is Shared,Primary SMTP, Mailbox Guid,DB Name, Mailbox Server, Mailbox Type,Mailbox Size
## create report for mailboxes with SendAs configured and Array of email addresses for the Migration batch csv
$MailBoxStatsReport = @()
$GroupMemberReport = @()
$SharedMailBoxReportData = @()
$NestedGroupReportData = @()
$MigrationBatchSharedReport = @()
$MigrationBatchShared = @()
$FullAccessGroupMemberReportData = @()
$FullAccessReportData = @()
$FullAccessNestedGroupReportData = @()
$MigrationBatchFullAccessReport = @()
$FullAccessMigrationBatchShared = @()
$Count = 0
$MailboxSendAsAccessReport = @()
$SendAsMigrationReport = @()
$SendAsBatch = @()
$GroupMemberSendOnBehalfReportData = @()
$SendOnBehalfMailBoxReportData = @()
$NestedGroupSendOnBehalfReportData = @()
$MigrationBatchSendOnBehalfReport = @()
$MigrationBatchSendOnBehalf = @()

foreach ($mailbox in $Mailboxes) {
    foreach ($MailBoxStat in (Get-MailboxStatistics $mailbox.UserPrincipalName)) {
        $ArchSize = $null
        $ArchSizeBytes = $null
		$Archive = Get-MailboxStatistics $mailbox.UserPrincipalName -archive -ErrorAction SilentlyContinue 
		If (!($Archive)) {
			Write-Host "$($MailBox.name) does not have an Archve mailbox" -foregroundcolor yellow
            }
        else{
            $ArchSize = $($Archive.TotalItemSize.Value)
            $ArchSizeBytes = $(($Archive.totalitemsize.value).ToString().Split(" ")[2].replace("(",''))
        }
         $MailBoxStatsReport += [PsCustomObject] [Ordered] @{
            "User" = "$($MailBoxStat.DisplayName)"
			"Account Disabled" = "$($mailbox.AccountDisabled)"
			"Alias" = "$($Mailbox.Alias)"
			"Is Resource" = "$($mailbox.IsResource)"
			"Is Shared" = "$($mailbox.IsShared)"
			"Primary SMTP" = "$($mailbox.PrimarySMTPAddress)"
			"Mailbox is Dirsynced" = "$($mailbox.IsDirSynced)"
			"Mailbox on Litigation Hold" = "$($MailBox.LitigationHoldEnabled)"
			"Mailbox move status" = "$($MailBox.MailboxMoveStatus)"
			"Mailbox Type" = "$($Mailbox.RecipientTypeDetails)"
			"Mailbox Created" = "$($Mailbox.WhenCreated)"
			"mailbox Guid" = "$($MailBoxStat.MailboxGuid)"
            "Database Name" = "$($MailBoxStat.DatabaseName)"
            "Server Name" ="$($MailBoxStat.ServerName)"
            "Mailbox Size" = "$($MailBoxStat.TotalItemSize.Value)"
			"SizeBytes" = "$(($MailBoxStat.totalitemsize.value).ToString().Split(" ")[2].replace("(",'')))"
			"Archive Size" = $ArchSize
			"Archive Size Bytes" = $ArchSizeBytes
			"LastLogonTime" = "$($MailBoxStat.LastLogonTime)"
			"RetentionPolicy" = "$($MailBox.RetentionPolicy)"
         }
	}
    $SendAs = Get-ADPermission -Identity $mailbox.Identity | Where-Object {($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and  ($_.User -notlike "NT AUTHORITY\SELF") -and ($_.user -notlike "S-1-*") }
    $Count++
    Write-host "Checking MAilbox $count of $($Mailboxes.count) " -ForegroundColor Green
    IF ($SendAs)  {
        foreach ($sender in $SendAs) {
            $User = Get-aduser -Identity $sender.user.RawIdentity.Split("\")[1] -Properties *
            $PrimaryEmail = $($mailbox.PrimarySmtpAddress.Local + "@" + $mailbox.PrimarySmtpAddress.Domain)
            Write-Host this account has sendAs access $SendAs.User to $mailbox.name -ForegroundColor Green
            Write-Host   $sender.User  has send as access to $PrimaryEmail mailbox name $mailbox.Name -ForegroundColor Green
            $SendAsBatch += $user.EmailAddress
            $SendAsBatch += $PrimaryEmail5
                $MailboxSendAsAccessReport += [PsCustomObject] [Ordered] @{
                    "MailboxName" = "$($mailbox.Name)"
                    "MailboxEmailAddress" = "$($PrimaryEmail)"
                    "MailBoxUPN" = "$($mailbox.UserPrincipalName)"
			        "MailboxPrimarySMTP" = "$($mailbox.PrimarySmtpAddress.ToString())"
			        "MailboxRecipientType" = "$($mailbox.RecipientType)"
			        "SendAsUser" ="$($sender.User.RawIdentity)"
			        "SendAsUseremail" = "$($User.EmailAddress)"
                    "SendAsUserEnabled" = "$($User.Enabled)"
                }
        }
    }
## Shared mailbox reports
	$FullAccessRights = Get-MailboxPermission -Identity $Mailbox.PrimarySmtpAddress.ToString() | Where-Object {($_.user -notlike "NT Authority*") -and !$_.IsInherited -and ($_.user -notlike "S-1-*") -and ($_.User -Notlike '*Discovery Management*') -and ($_.User -Notlike '*Organization Management*')}
	#No non-default permissions found, continue to next mailbox
	if (!$FullAccessRights) { continue }
		foreach ($Entry in $FullAccessRights) {
			$AccessAccount = $null
			Try {
				$AccessAccount = Get-ADUser $Entry.user.SecurityIdentifier.value -Properties *
			}
			Catch {
			Try {
				$AccessAccount = Get-ADGroup $Entry.user.SecurityIdentifier.value
			}
			Catch {write-host $AccessAccount.Name is not a user or group -ForegroundColor Red}
			}
			If ($Mailbox.RecipientTypeDetails -eq "SharedMailbox") {
				If ($AccessAccount.objectClass -eq "group") {
					foreach ($GroupMember in Get-ADGroupMember -Identity $AccessAccount.SID.Value) {
						if ($GroupMember.ObjectClass -eq "user") {
							$GroupUser = Get-ADUser -identity $GroupMember.SID.Value -Properties *
						}
						If ($GroupMember.objectClass -eq "group") {
							$NestedGroupReportData += [PsCustomObject] [Ordered] @{
							"Parent Group SAM name" = "$($AccessAccount.SamAccountName)"
							"Parent Group Category" = "$($AccessAccount.GroupCategory)"
							"Parent Group Scope" = "$($AccessAccount.GroupScope)"
							"Parent Group Class" = "$($AccessAccount.ObjectClass)"
							"Parent Group SID" = "$($AccessAccount.SID.Value)"
							"Nested Group SAM Name" = "$($GroupMember.SamAccountName)"
							"Nested Group Class" = "$($GroupMember.objectClass)"
							"Nested Group SID" = "$($GroupMember.SID.Value)"
							}
						}
						$GroupMemberReportData += [PsCustomObject] [Ordered] @{
						"Mailbox name" = "$($Mailbox.Name)"
						"Mailbox recipient type" = "$($Mailbox.RecipientTypeDetails)"
						"Mailbox Alias" = "$($Mailbox.Alias)"
						"Mailbox Database" = "$($Mailbox.Database.Rdn.EscapedName)"
						"Mailbox Server" = "$($Mailbox.ServerName)"
						"Shared Mailbox Primary SMTP" = "$($Mailbox.PrimarySmtpAddress.ToString())"
						"Group Name" = "$($AccessAccount.Name)"
						"Group Member" = "$($GroupMember.name)"
						"Group Member Obj Type" = "$($GroupMember.objectClass)"
						"Group User EmailAddress" = "$($GroupUser.EmailAddress)"
						"Group Member Sam Account" = "$($GroupMember.SamAccountName)"
						"Group Member UPN" = "$($GroupUser.UserPrincipalName)"
						"Group Member Access Type" = "$($Entry.AccessRights -join ",")"
						}
					}
				}
			$SharedMailBoxReportData += [PsCustomObject] [Ordered] @{
			"Mailbox name" = "$($Mailbox.Name)"
			"Mailbox recipient type" = "$($Mailbox.RecipientTypeDetails)"
			"Mailbox Alias" = "$($Mailbox.Alias)"
			"Mailbox Database" = "$($Mailbox.Database.Rdn.EscapedName)"
			"Mailbox Server" = "$($Mailbox.ServerName)"
			"Shared Mailbox Primary SMTP" = "$($Mailbox.PrimarySmtpAddress.toString())"
			"Account with Access" = "$($Entry.user)"
			"Access Account Obj Type" = "$($AccessAccount.ObjectClass)"
			"Access account SAM" = "$($AccessAccount.SamAccountName)"
			"Access Account UPN" = "$($AccessAccount.UserPrincipalName)"
			"Access Account Email Address" = "$($AccessAccount.Emailaddress)"
			"Access Rights" = "$($Entry.AccessRights -join ",")"
			}
			$MigrationBatchShared += $Mailbox.PrimarySmtpAddress.ToString()
			$MigrationBatchShared += $AccessAccount.EmailAddress
			}
			If ($AccessAccount.objectClass -eq "group") {
				foreach ($GroupMember in Get-ADGroupMember -Identity $AccessAccount.SID.Value) {
					if ($GroupMember.ObjectClass -eq "user") {
						$GroupUser = Get-ADUser -identity $GroupMember.SID.Value -Properties *
					}
					If ($GroupMember.objectClass -eq "group") {
						$FullAccessNestedGroupReportData += [PsCustomObject] [Ordered] @{
						"Parent Group SAM name" = "$($AccessAccount.SamAccountName)"
						"Parent Group Category" = "$($AccessAccount.GroupCategory)"
						"Parent Group Scope" = "$($AccessAccount.GroupScope)"
						"Parent Group Class" = "$($AccessAccount.ObjectClass)"
						"Parent Group SID" = "$($AccessAccount.SID.Value)"
						"Nested Group SAM Name" = "$($GroupMember.SamAccountName)"
						"Nested Group Class" = "$($GroupMember.objectClass)"
						"Nested Group SID" = "$($GroupMember.SID.Value)"
						}
					}
					$FullAccessGroupMemberReportData += [PsCustomObject] [Ordered] @{
					"Mailbox name" = "$($Mailbox.Name)"
					"Mailbox recipient type" = "$($Mailbox.RecipientTypeDetails)"
					"Mailbox Alias" = "$($Mailbox.Alias)"
					"Mailbox Database" = "$($Mailbox.Database.Rdn.EscapedName)"
					"Mailbox Server" = "$($Mailbox.ServerName)"
					"Shared Mailbox Primary SMTP" = "$($Mailbox.PrimarySmtpAddress.ToString())"
					"Group Name" = "$($AccessAccount.Name)"
					"Group Member" = "$($GroupMember.name)"
					"Group Member Obj Type" = "$($GroupMember.objectClass)"
					"Group User EmailAddress" = "$($GroupUser.EmailAddress)"
					"Group Member Sam Account" = "$($GroupMember.SamAccountName)"
					"Group Member UPN" = "$($GroupUser.UserPrincipalName)"
					"Group Member Access Type" = "$($Entry.AccessRights -join ",")"
					"Access Deny" = "$($Entry.Deny)"
					}
				}
			}
		$FullAccessReportData += [PsCustomObject] [Ordered] @{
		"Mailbox name" = "$($Mailbox.Name)"
		"Mailbox recipient type" = "$($Mailbox.RecipientTypeDetails)"
		"Mailbox Alias" = "$($Mailbox.Alias)"
		"Mailbox Database" = "$($Mailbox.Database.Rdn.EscapedName)"
		"Mailbox Server" = "$($Mailbox.ServerName)"
		"Shared Mailbox Primary SMTP" = "$($Mailbox.PrimarySmtpAddress.toString())"
		"Account with Access" = "$($Entry.user)"
		"Access Account Obj Type" = "$($AccessAccount.ObjectClass)"
		"Access account SAM" = "$($AccessAccount.SamAccountName)"
		"Access Account UPN" = "$($AccessAccount.UserPrincipalName)"
		"Access Account Email Address" = "$($AccessAccount.Emailaddress)"
		"Access Rights" = "$($Entry.AccessRights -join ",")"
		"Access Deny" = "$($Entry.Deny)"
		}
		$FullAccessMigrationBatchShared += $Mailbox.PrimarySmtpAddress.ToString()
		$FullAccessMigrationBatchShared += $AccessAccount.EmailAddress
	}
## Grant Send On Behalf Of reports
	If ($Mailbox.GrantSendOnBehalfTo) {
		foreach ($Entry in $Mailbox.GrantSendOnBehalfTo) {
			$AccessAccount = $null
			Try {
				$AccessAccount = Get-ADUser $Entry.DistinguishedName -Properties *
			}
			Catch {
			Try {
				$AccessAccount = Get-ADGroup $Entry.DistinguishedName
			}
			Catch {write-host $AccessAccount.Name is not a user or group -ForegroundColor Red}
			}
			If ($AccessAccount.objectClass -eq "group") {
				foreach ($GroupMember in Get-ADGroupMember -Identity $AccessAccount.SID.Value) {
					if ($GroupMember.ObjectClass -eq "user") {
						$GroupUser = Get-ADUser -identity $GroupMember.SID.Value -Properties *
					}
					If ($GroupMember.objectClass -eq "group") {
						$NestedGroupSendOnBehalfReportData += [PsCustomObject] [Ordered] @{
						"Parent Group SAM name" = "$($AccessAccount.SamAccountName)"
						"Parent Group Category" = "$($AccessAccount.GroupCategory)"
						"Parent Group Scope" = "$($AccessAccount.GroupScope)"
						"Parent Group Class" = "$($AccessAccount.ObjectClass)"
						"Parent Group SID" = "$($AccessAccount.SID.Value)"
						"Nested Group SAM Name" = "$($GroupMember.SamAccountName)"
						"Nested Group Class" = "$($GroupMember.objectClass)"
						"Nested Group SID" = "$($GroupMember.SID.Value)"
						}
					}
					$GroupMemberSendOnBehalfReportData += [PsCustomObject] [Ordered] @{
					"Mailbox name" = "$($Mailbox.Name)"
					"Mailbox Alias" = "$($Mailbox.Alias)"
					"Mailbox Database" = "$($Mailbox.Database.Rdn.EscapedName)"
					"Mailbox Server" = "$($Mailbox.ServerName)"
					"Mailbox Primary SMTP" = "$($Mailbox.PrimarySmtpAddress.ToString())"
					"Group Name" = "$($AccessAccount.Name)"
					"Group Member" = "$($GroupMember.name)"
					"Group Member Obj Type" = "$($GroupMember.objectClass)"
					"Group User EmailAddress" = "$($GroupUser.EmailAddress)"
					"Group Member Sam Account" = "$($GroupMember.SamAccountName)"
					"Group Member UPN" = "$($GroupUser.UserPrincipalName)"
					"Group Member Access Type" = "$($Entry.AccessRights -join ",")"
					}
				}
			}
			$SendOnBehalfMailBoxReportData += [PsCustomObject] [Ordered] @{
			"Mailboxname" = "$($Mailbox.Name)"
			"MailboxAlias" = "$($Mailbox.Alias)"
			"MailboxDatabase" = "$($Mailbox.Database.Rdn.EscapedName)"
			"MailboxServer" = "$($Mailbox.ServerName)"
			"MailboxPrimarySMTP" = "$($Mailbox.PrimarySmtpAddress.toString())"
			"AccountwithAccess" = "$($Entry.user)"
			"AccessAccountObjectClass" = "$($AccessAccount.ObjectClass)"
			"AccessAccountSAM" = "$($AccessAccount.SamAccountName)"
			"AccessAccountUPN" = "$($AccessAccount.UserPrincipalName)"
			"AccessAccountEmailAddress" = "$($AccessAccount.Emailaddress)"
			"AccessRights" = "$($Entry.AccessRights -join ",")"
			}
			$MigrationBatchSendOnBehalf += $Mailbox.PrimarySmtpAddress.ToString()
			$MigrationBatchSendOnBehalf += $AccessAccount.EmailAddress
		}
	}
}

$FullAccessNestedGroupReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$FullAccessNestedGroupReportName"
$FullAccessGroupMemberReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$FullAccessGroupMemberReportName"
$FullAccessReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$FullAccessReportName"
$NestedGroupReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$NestedGroupReportName"
$GroupMemberReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$ShareMailBoxGroupMemberReportName"
$SharedMailBoxReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$SharedMailBoxReportName"
$NestedGroupSendOnBehalfReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$NestedGroupSendOnBehalfReportName"
$GroupMemberSendOnBehalfReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$MailboxGroupMemberSendOnBehalfReportName"
$SendOnBehalfMailBoxReportData | Export-Csv -NoTypeInformation -Path "$ReportPath\$MailboxSendOnBehalfReportName"
$MailBoxStatsReport | Export-Csv -NoTypeInformation "$ReportPath\$MailBoxStatsReportName"
$MailboxSendAsAccessReport | Export-Csv -NoTypeInformation "$ReportPath\$MailboxSendAsAccessReportName"
## Create Mirgration batch CSV for mailboxes with Full permissions
ForEach ($MBShare in $MigrationBatchShared | Select-Object -Unique) {
    $MigrationBatchSharedReport += [PsCustomObject] [Ordered] @{
        "emailaddress" = $MBShare
    }
}
$MigrationBatchSharedReport | Export-Csv -NoTypeInformation "$ReportPath\$MigrationBatchSharedReportName"

## Create Mirgration batch CSV for shared mailboxes
ForEach ($MBShare in $FullAccessMigrationBatchShared | Select-Object -Unique) {
    $MigrationBatchFullAccessReport += [PsCustomObject] [Ordered] @{
        "emailaddress" = $MBShare
    }
}
$MigrationBatchFullAccessReport | Export-Csv -NoTypeInformation "$ReportPath\$MigrationBatchFullAccessReportName"

## Create Mirgration batch CSV for mailboxes with sendAS permissions
ForEach ($SendAs in $SendAsBatch | Select-Object -Unique) {
    $SendAsMigrationReport += [PsCustomObject] [Ordered] @{
        "emailaddress" = $SendAs
    }
}
$SendAsMigrationReport | Export-Csv -NoTypeInformation "$ReportPath\$MigrationBatchSendAsName"

## Create Mirgration batch CSV for mailboxes with SenOnBehalf permissions
ForEach ($MBShare in $MigrationBatchSendOnBehalf | Select-Object -Unique) {
    $MigrationBatchSendOnBehalfReport += [PsCustomObject] [Ordered] @{
        "emailaddress" = $MBShare
    }
}
$MigrationBatchSendOnBehalfReport | Export-Csv -NoTypeInformation "$ReportPath\$MigrationBatchSendOnBehalfReportName"

### The following will list all smtp addresses for user mailboxes  ###
$Mailboxes |Select-Object DisplayName,ServerName,PrimarySmtpAddress, @{Name="EmailAddresses";Expression={$_.EmailAddresses |Where-Object {$_.PrefixString -ceq "smtp"} | ForEach-Object {$_.SmtpAddress}}} | Export-csv "$ReportPath\$SmtpAddressReportName" -NoTypeInformation

### The following scripts output any forwarders configured on mailboxes ###
$Mailboxes | Where-Object {($_.ForwardingAddress -ne $Null) -or ($_.ForwardingsmtpAddress -ne $Null)} | Select-Object Name, DisplayName, PrimarySMTPAddress, UserPrincipalName, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward | Export-csv -NoTypeInformation "$ReportPath\$MailboxesWithForwarding"


########################################################

### The following scripts output mailbox statistics ###
$Mailboxes | Group-object recipienttypedetails | Select-Object count, name | Out-File "$ReportPath\$MailboxCountName"

## This is here for legacy systems show servers that host autodiscover and Outlook anywhere there is also a csv in this script using the newer command
Get-ClientAccessServer | Select-Object  Name,AutoDiscoverServiceCN,AutoDiscoverServiceInternalUri,OutlookAnywhereEnabled | Format-List | Out-File "$ReportPath\$ClientAccessServerTXT"

## Email Address Policy report
Get-EmailAddressPolicy | Select-Object Name,Priority,RecipientFilter,RecipientFilterApplied,IncludeRecipients,EnabledPrimarySMTPAddressTemplate,EnabledEmailAddressTemplates,Enabled,IsValid | Out-File "$ReportPath\$EmailAddressPolicy"

## Transport  Configuration report
Get-TransportServer | Select-Object Name,InternalDNSServers,ExternalDNSServers,OutboundConnectionFailureRetryInterval,TransientFailureRetryInterval,TransientFailureRetryCount,MessageExpirationTimeout,DelayNotificationTimeout,MaxOutboundConnections,MaxPerDomainOutboundConnections,MessageTrackingLogEnabled,MessageTrackingLogPath,ConnectivityLogEnabled,ConnectivityLogPath,SendProtocolLogPath,ReceiveProtocolLogPath | Out-File "$ReportPath\$TransportConfiguration"

## OWA Policy report
Get-OwaMailboxPolicy | Select-Object Name,ActiveSyncIntegrationEnabled,AllAddressListsEnabled,CalendarEnabled,ContactsEnabled,JournalEnabled,JunkEmailEnabled,RemindersAndNotificationsEnabled,NotesEnabled,PremiumClientEnabled,SearchFoldersEnabled,SignaturesEnabled,SpellCheckerEnabled,TasksEnabled,ThemeSelectionEnabled,UMIntegrationEnabled,ChangePasswordEnabled,RulesEnabled,PublicFoldersEnabled,SMimeEnabled,RecoverDeletedItemsEnabled,InstantMessagingEnabled,TextMessagingEnabled,DirectFileAccessOnPublicComputersEnabled,WebReadyDocumentViewingOnPublicComputersEnabled,DirectFileAccessOnPrivateComputersEnabled,WebReadyDocumentViewingOnPrivateComputersEnabled | Out-File "$ReportPath\$OWAPolicies"

## ActiveSync Mobile Device Policy report
Get-ActiveSyncMailboxPolicy | Select-Object Name,AllowNonProvisionableDevices,DevicePolicyRefreshInterval,PasswordEnabled,MaxCalendarAgeFilter,MaxEmailAgeFilter,MaxAttachmentSize,RequireManualSyncWhenRoaming,AllowHTMLEmail,AttachmentsEnabled,AllowStorageCard,AllowCameraTrue,AllowWiFi,AllowIrDA,AllowInternetSharing,AllowRemoteDesktop,AllowDesktopSync,AllowBluetooth,AllowBrowser,AllowConsumerEmail,AllowUnsignedApplications,AllowUnsignedInstallationPackages | Out-File "$ReportPath\$MobileDevicePolicy"

## Transport rule report
Get-TransportRule | Format-List Name,Priority,Description,SenderIpRanges,Comments,State | Out-File "$ReportPath\$TransportRules"

### Get Virtual Directory Reports ###
$OwaVirtDirReport = @()
$EcpVirtDirReport = @()
$ActSyncVirtDirReport = @()
$EwsVirtDirReport = @()
$OabVirtDirReport = @()
$MapiVirtDirReport = @()
$oAnyVirtDirReport = @()
$AutoDiscVirtDirReport = @()
$PsVirtDirReport = @()
foreach ($server in $Servers) {
	$Owa = $null
	$OwaIntDNS = $null
	$OwaExtDNS = $null
    $Owa = Get-OwaVirtualDirectory -Server $server.name
	if ($Owa.InternalUrl) {
		$OwaIntDNS = (Resolve-DnsName -Name $($Owa.InternalUrl.AbsoluteUri.Split("/")[2]) -Type A -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($Owa.ExternalUrl) {
		$OwaExtDNS = (Resolve-DnsName -Name $($Owa.ExternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$OwaVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"Owa Name" = "$($Owa.Name)"
		"Owa Internal Url" = "$($Owa.InternalUrl)"
		"Resolved IP Int URL" = $OwaIntDNS
		"Owa External Url" ="$($Owa.ExternalUrl)"
		"Resolved IP Ext URL" = $OwaExtDNS
		"Owa Fail back Url" ="$($Owa.Failbackurl)"
		"Owa Adfs Authentication" = "$($Owa.AdfsAuthentication)"
		"Owa Basic Authentication" = "$($Owa.BasicAuthentication)"
		"Owa Forms Authentication" = "$($Owa.FormsAuthentication)"
		"Owa OAuth Authentication" = "$($Owa.OAuthAuthentication)"
		"Owa Windows Authentication" = "$($Owa.WindowsAuthentication)"
		"Owa Internal Auth Methods" = "$($Owa.InternalAuthenticationMethods)"
		"Owa External Auth Methods" = "$($Owa.ExternalAuthenticationMethods)"
	}
	$Ecp = $null
	$ECPIntDNS = $null
	$ECPExtDNS = $null
    $Ecp = Get-EcpVirtualDirectory -Server $server.name
	if ($ECP.InternalUrl) {
		$ECPIntDNS = (Resolve-DnsName -Name $($ECP.InternalUrl.AbsoluteUri.Split("/")[2]) -Type A -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($ECP.ExternalUrl) {
		$ECPExtDNS = (Resolve-DnsName -Name $($ECP.ExternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$EcpVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"Ecp Name" = "$($Ecp.Name)"
		"Ecp Internal Url" = "$($Ecp.InternalUrl)"
		"Resolved IP Int URL" = $EcpIntDNS
		"Ecp External Url" ="$($Ecp.ExternalUrl)"
		"Resolved IP Ext URL" = $EcpExtDNS
		"Ecp Adfs Authentication" = "$($Ecp.AdfsAuthentication)"
		"Ecp Basic Authentication" = "$($Ecp.BasicAuthentication)"
		"Ecp Forms Authentication" = "$($Ecp.FormsAuthentication)"
		"Ecp OAuth Authentication" = "$($Ecp.OAuthAuthentication)"
		"Ecp Windows Authentication" = "$($Ecp.WindowsAuthentication)"
		"Ecp Internal Auth Methods" = "$($Ecp.InternalAuthenticationMethods)"
		"Ecp External Auth Methods" = "$($Ecp.ExternalAuthenticationMethods)"
	}
	$ActSync = $null
	$ActSyncIntDNS = $null
	$ActSyncExtDNS = $null
	$actSync = Get-ActiveSyncVirtualDirectory -Server $server.name
	if ($ActSync.InternalUrl) {
		$ActSyncIntDNS = (Resolve-DnsName -Name $($ActSync.InternalUrl.AbsoluteUri.Split("/")[2]) -Type A -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($ActSync.ExternalUrl) {
		$ActSyncExtDNS = (Resolve-DnsName -Name $($ActSync.ExternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$ActSyncVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"ActSync Name" = "$($ActSync.Name)"
		"ActSync Internal Url" = "$($ActSync.InternalUrl)"
		"Resolved IP Int URL" = $ActSyncIntDNS
		"ActSync External Url" ="$($ActSync.ExternalUrl)"
		"Resolved IP Ext URL" = $ActSyncExtDNS
		"ActSync Basic Authentication" = "$($ActSync.BasicAuthEnabled)"
		"ActSync Windows Authentication" = "$($ActSync.WindowsAuthEnabled)"
        "ActSync Virtual Dir Name" = "$($actSync.VirtualDirectoryName)"
        "ActSync SSL Enabled" = "$($actSync.WebSiteSSLEnabled)"
	}
	$Ews = $null
	$EwsIntDNS = $null
	$EwsExtDNS = $null
    $Ews = Get-WebServicesVirtualDirectory -Server $server.name
	if ($Ews.InternalUrl) {
		$EwsIntDNS = (Resolve-DnsName -Name $($Ews.InternalUrl.AbsoluteUri.Split("/")[2]) -Type A -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($Ews.ExternalUrl) {
		$EwsExtDNS = (Resolve-DnsName -Name $($Ews.ExternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$EwsVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"Ews Name" = "$($Ews.Name)"
		"Ews Internal Url" = "$($Ews.InternalUrl)"
		"Resolved IP Int URL" = $EwsIntDNS
		"Ews External Url" ="$($Ews.ExternalUrl)"
		"Resolved IP Ext URL" = $EwsExtDNS
		"Ews Adfs Authentication" = "$($Ews.AdfsAuthentication)"
		"Ews Basic Authentication" = "$($Ews.BasicAuthentication)"
		"Ews Certificate Authentication" = "$($Ews.CertificateAuthentication)"
		"Ews Digest Authentication" = "$($Ews.DigestAuthentication)"
		"Ews LiveID Basic Authentication" = "$($Ews.LiveIdBasicAuthentication)"
		"Ews LiveID Negotiate Authentication" = "$($Ews.LiveIdNegotiateAuthentication)"
		"Ews OAuth Authentication" = "$($Ews.OAuthAuthentication)"
		"Ews Windows Authentication" = "$($Ews.WindowsAuthentication)"
		"Ews WS Security Authentication" = "$($Ews.WSSecurityAuthentication)"
        "Ews Internal Auth Methods" = "$($Ews.InternalAuthenticationMethods)"
        "Ews External Auth Methods" = "$($Ews.ExternalAuthenticationMethods)"
		"MRSProxyEnabled" = "$($Ews.MRSProxyEnabled)"
	}
	$Oab = $null
	$OabIntDNS = $null
	$OabExtDNS = $null
    $Oab = Get-OabVirtualDirectory -Server $server.name
	if ($Oab.InternalUrl) {
		$OabIntDNS = (Resolve-DnsName -Name $($Oab.InternalUrl.AbsoluteUri.Split("/")[2]) -Type A -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($Oab.ExternalUrl) {
		$OabExtDNS = (Resolve-DnsName -Name $($Oab.ExternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$OabVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"Oab Name" = "$($Oab.Name)"
		"Oab Internal Url" = "$($Oab.InternalUrl)"
		"Resolved IP Int URL" = $OabIntDNS
		"Oab External Url" ="$($Oab.ExternalUrl)"
		"Resolved IP Ext URL" = $OabExtDNS
		"Oab Basic Authentication" = "$($Oab.BasicAuthentication)"
		"Oab OAuth Authentication" = "$($Oab.OAuthAuthentication)"
		"Oab Windows Authentication" = "$($Oab.WindowsAuthentication)"
        "Oab Internal Auth Methods" = "$($Oab.InternalAuthenticationMethods)"
        "Oab External Auth Methods" = "$($Oab.ExternalAuthenticationMethods)"
		"Oab Require SSL" = "$($Oab.RequireSSL)"
	}
	$Mapi = $null
	$MapiIntDNS = $null
	$MapiExtDNS = $null
    $Mapi = Get-MapiVirtualDirectory -Server $server.name
	if ($Mapi.InternalUrl) {
		$MapiIntDNS = (Resolve-DnsName -Name $($Mapi.InternalUrl.AbsoluteUri.Split("/")[2]) -Type A -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($Mapi.ExternalUrl) {
		$MapiExtDNS = (Resolve-DnsName -Name $($Mapi.ExternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$MapiVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"Mapi Name" = "$($Mapi.Name)"
		"Mapi Internal Url" = "$($Mapi.InternalUrl)"
		"Resolved IP Int URL" = $MapiIntDNS
		"Mapi External Url" ="$($Mapi.ExternalUrl)"
		"Resolved IP Ext URL" = $MapiExtDNS
        "Mapi Internal Auth Methods" = "$($Mapi.InternalAuthenticationMethods)"
        "Mapi External Auth Methods" = "$($Mapi.ExternalAuthenticationMethods)"
		"Mapi IIS Auth Methods" = "$($Mapi.IISAuthenticationMethods)"
	}
	$oAny = $null
	$oAnyIntDNS = $null
	$oAnyExtDNS = $null
    $oAny = Get-OutlookAnywhere -Server $server.name
	if ($oAny.InternalHostname) {
		$oAnyIntDNS = (Resolve-DnsName -Name $($oAny.InternalHostname) -Type A -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($oAny.ExternalHostname) {
		$oAnyExtDNS = (Resolve-DnsName -Name $($oAny.ExternalHostname) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$oAnyVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"oAny Name" = "$($oAny.Name)"
		"oAny Internal Url" = "$($oAny.InternalHostname)"
		"Resolved IP Int URL" = $oAnyIntDNS
		"oAny External Url" ="$($oAny.ExternalHostname)"
		"Resolved IP Ext URL" = $oAnyExtDNS
        "oAny Internal Auth Methods" = "$($oAny.InternalClientAuthenticationMethod)"
        "oAny External Auth Methods" = "$($oAny.ExternalClientAuthenticationMethod)"
		"oAny IIS Auth Methods" = "$($oAny.IISAuthenticationMethods)"
		"oAny Require SSL Internal" = "$($oAny.InternalClientsRequireSsl)"
		"oAny Require SSL External" = "$($oAny.ExternalClientsRequireSsl)"
	}
	$AutoDicoverVirtDir = $null
	$AutoDicoverVirtDirIntDNS = $null
	$AutoDicoverVirtDirExtDNS = $null
	$AutoDiscoverIntDNSType = $null
	$AutoDiscoverExtDNSType = $null
	$AutoDicoverVirtDir = Get-AutodiscoverVirtualDirectory -Server $server.name
	if ($AutoDicoverVirtDir.InternalUrl) {
		$AutoDiscoverIntDNSType = (Resolve-DnsName -Name $($AutoDicoverVirtDir.InternalUrl.AbsoluteUri.Split("/")[2]) -Type All -ErrorAction SilentlyContinue).QueryType
		$AutoDicoverVirtDirIntDNS = (Resolve-DnsName -Name $($AutoDicoverVirtDir.InternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($AutoDicoverVirtDir.ExternalUrl) {
		$AutoDiscoverExtDNSType = (Resolve-DnsName -Name $($AutoDicoverVirtDir.InternalUrl.AbsoluteUri.Split("/")[2]) -Type All -ErrorAction SilentlyContinue).QueryType
		$AutoDicoverVirtDirExtDNS = (Resolve-DnsName -Name $($AutoDicoverVirtDir.ExternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$AutoDiscVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"AutoDiscover Name" = "$($AutoDicoverVirtDir.Name)"
		"AutoDiscover Internal Url" = "$($AutoDicoverVirtDir.InternalUrl)"
		"Resolved IP Int URL" = $AutoDicoverVirtDirIntDNS
		"Int Record Type" = $AutoDiscoverIntDNSType
		"AutoDiscover External Url" ="$($AutoDicoverVirtDir.ExternalUrl)"
		"Resolved IP Ext URL" = $AutoDicoverVirtDirExtDNS
		"Ext Record Type" = $AutoDiscoverExtDNSType
		"AutoDiscover Adfs Authentication" = "$($AutoDicoverVirtDir.AdfsAuthentication)"
		"AutoDiscover Basic Authentication" = "$($AutoDicoverVirtDir.BasicAuthentication)"
		"AutoDiscover Digest Authentication" = "$($AutoDicoverVirtDir.DigestAuthentication)"
		"AutoDiscover LiveID Basic Authentication" = "$($AutoDicoverVirtDir.LiveIdBasicAuthentication)"
		"AutoDiscover LiveID Negotiate Authentication" = "$($AutoDicoverVirtDir.LiveIdNegotiateAuthentication)"
		"AutoDiscover OAuth Authentication" = "$($AutoDicoverVirtDir.OAuthAuthentication)"
		"AutoDiscover Windows Authentication" = "$($AutoDicoverVirtDir.WindowsAuthentication)"
		"AutoDiscover WS Security Authentication" = "$($AutoDicoverVirtDir.WSSecurityAuthentication)"
		"AutoDiscover Internal Auth Methods" = "$($AutoDicoverVirtDir.InternalAuthenticationMethods)"
		"AutoDiscover External Auth Methods" = "$($AutoDicoverVirtDir.ExternalAuthenticationMethods)"
	}
	$PsVdir = $null
	$PsVDirIntDNS = $null
	$PsVDirExtDNS = $null
	$PsVdir = Get-PowerShellVirtualDirectory -Server $server.Name
	if ($PsVdir.InternalUrl) {
		$PsVDirIntDNS = (Resolve-DnsName -Name $($PsVdir.InternalUrl.AbsoluteUri.Split("/")[2]) -Type A -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	if ($PsVdir.ExternalUrl) {
		$PsVDirExtDNS = (Resolve-DnsName -Name $($PsVdir.ExternalUrl.AbsoluteUri.Split("/")[2]) -Type A -Server $PublicDNSServer -ErrorAction SilentlyContinue).IPAddress -join ", "
	}
	$PsVirtDirReport += [PsCustomObject] [Ordered] @{
		"Server Name" = "$($Server.Name)"
		"PsVDir Name" = "$($PsVDir.Name)"
		"PsVDir Internal Url" = "$($PsVDir.InternalUrl)"
		"Resolved IP Int URL" = $PsVDirIntDNS
		"PsVDir External Url" ="$($PsVDir.ExternalUrl)"
		"Resolved IP Ext URL" = $PsVDirExtDNS
		"PsVDir Adfs Authentication" = "$($PsVDir.AdfsAuthentication)"
		"PsVDir Basic Authentication" = "$($PsVDir.BasicAuthentication)"
		"PsVDir Certificate Authentication" = "$($PsVDir.CertificateAuthentication)"
		"PsVDir Digest Authentication" = "$($PsVDir.DigestAuthentication)"
		"PsVDir LiveID Basic Authentication" = "$($PsVDir.LiveIdBasicAuthentication)"
		"PsVDir LiveID Negotiate Authentication" = "$($PsVDir.LiveIdNegotiateAuthentication)"
		"PsVDir OAuth Authentication" = "$($PsVDir.OAuthAuthentication)"
		"PsVDir Windows Authentication" = "$($PsVDir.WindowsAuthentication)"
		"PsVDir WS Security Authentication" = "$($PsVDir.WSSecurityAuthentication)"
        "PsVDir Internal Auth Methods" = "$($PsVDir.InternalAuthenticationMethods)"
        "PsVDir External Auth Methods" = "$($PsVDir.ExternalAuthenticationMethods)"
		"PsVDir Require SSL" = "$($PsVDir.RequireSSL)"
    }
}
$OwaVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$OwaVirtDirReportName"
$EcpVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$EcpVirtDirReportName"
$ActSyncVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$ActSyncVirtDirReportName"
$EwsVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$EwsVirtDirReportName"
$OabVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$OabVirtDirReportName"
$MapiVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$MapiVirtDirReportName"
$oAnyVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$oAnyVirtDirReportName"
$AutoDiscVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$AutoDiscVirtDirReportName"
$PsVirtDirReport | Export-Csv -NoTypeInformation -Path "$ReportPath\$PsVirtDirReportName"



