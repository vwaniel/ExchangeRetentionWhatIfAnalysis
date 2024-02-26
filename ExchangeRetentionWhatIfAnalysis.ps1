<#
.SYNOPSIS
	This script reports on item size/count for mail older than the specified date in each user's mailbox.
.DESCRIPTION
	This script reports on item size/count for mail older than the specified date in each user's mailbox.
.PARAMETER SearchDate
	The cutoff date for items that will be counted.  Items newer than the specified date will not be counted.
.PARAMETER Credential
	The credentials that will be used to connect to Exchange and Exchange Web Services.  Requires mailbox impersionation permissions.  If credentials are not specified as a parameter the user will be prompted to enter them.
.PARAMETER ExchangeHost
	The hostname of the Exchange server that will be used to access the Exchange Management Shell via PS Remoting.
.PARAMETER EWSURI
	The URI used to access Exchange Web Services.
.PARAMETER MailboxName
	Optional parameter to specify the identity (or a comma-separated list of identities) of mailboxes to search.  If MailboxName is not specified, all user mailboxes will be searched.
.EXAMPLE
	PS> .\retention_limit_whatif_report.ps1 -SearchDate 1/1/2020 -ExchangeHost exch01.company.com -EWSURI https://exch.company.com/EWS/Exchange.asmx
	Searches all mailboxes for mail items older than January 1, 2020.
.EXAMPLE
	PS> .\retention_limit_whatif_report.ps1 -SearchDate 1/1/2020 -MailboxName test_user -ExchangeHost exch01.company.com -EWSURI https://exch.company.com/EWS/Exchange.asmx
	Searches the test_user mailbox (if it exists) for mail items older than January 1, 2020.
#>

#Requires -Module ActiveDirectory

[cmdletbinding()]
param(
	[Parameter(Mandatory=$true,Position=1)][datetime]$SearchDate,
	[Parameter(Mandatory=$false)][pscredential]$Credential,
	[Parameter(Mandatory=$true)][string]$ExchangeHost,
	[Parameter(Mandatory=$true)][string]$EWSURI,
	[Parameter(Mandatory=$false)][string[]]$MailboxName
)

# Get credentials if they haven't been provided.
if (-NOT $Credential) {
	$Credential = $host.UI.PromptForCredential("Exchange Credentials","Please enter the username and password with rights to Exchange.","","")
	if (-NOT $Credential) {
		Write-Error "Credentials must be specified to connect to Exchange."
		exit
	}
}

# Build KQL query.
$strKQLQuery = "(Received<$(Get-Date $SearchDate -format 'yyyy-MM-dd')) AND (Kind:email)"

# Establish PS Session
try {
	$objPSSession = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionURI "http://$($ExchangeHost)/PowerShell" -Authentication "Kerberos" -Credential $Credential -ErrorAction "Stop"
}
catch {
	Write-Error "Unable to connect to Exchange Server $($ExchangeHost).  $($_.Exception.Message)"
	exit
}

$objPSModule = Import-PSSession $objPSSession

$arrMailboxes = [System.Collections.ArrayList]@()
if ($MailboxName) {
	foreach ($name in $MailboxName) {
		Get-Mailbox -Identity $name -RecipientTypeDetails UserMailbox | Foreach-Object {$arrMailboxes.Add($_) > $null}
	}
} else {
	Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Foreach-Object {$arrMailboxes.Add($_) > $null}
}

# Exchange Web Services Searching Foo
if (Test-Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll") {
	try {
		Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" -ErrorAction "Stop"
	}
	catch {
		Write-Error "Unable to load Exchange Web Services DLL.  $($_.Exception.Message)"
		exit
	}
}

Remove-Module $objPSModule > $null

$objEWSService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013))
$objEWSService.Credentials = $Credential.GetNetworkCredential()
$objEWSService.URL = [System.URI]$EWSURI

$arrResults = [System.Collections.ArrayList]@()
foreach ($objMailbox in $arrMailboxes) {
	Write-Progress -Activity "Searching Mailboxes" -Status "$($objMailbox.PrimarySMTPAddress)" -PercentComplete (($($arrMailboxes.IndexOf($objMailbox))/$($arrMailboxes | Measure-Object).Count) * 100) -Id 1
	$objMBResponse = $objEWSService.GetSearchableMailboxes($objMailbox.PrimarySMTPAddress,$false)
	$objMBScope = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $objMBResponse.SearchableMailboxes.Length
	$objMBResponse.SearchableMailboxes | Foreach-Object {
		$objMBScope[$objMBResponse.SearchableMailboxes.IndexOf($_)] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($_.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All)
	}
	$objMailboxQuery = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($strKQLQuery,$objMBScope)
	$objSearchMailboxParameters = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters -Property @{
		"SearchQueries" = $objMailboxQuery;
		"PageSize" = 100;
		"PageDirection" = $([Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next);
		"PerformDeduplication" = $false;
		"ResultType" = $([Microsoft.Exchange.WebServices.Data.SearchResultType]::PreviewOnly)
	}
	$objSearchResults = $objEWSService.SearchMailboxes($objSearchMailboxParameters)
	if ($objSearchResults) {
		# Get the AD user account.
		$objADUser = Get-ADUser -Identity $objMailbox.SamAccountName -Properties "displayName","enabled","l","department","title","mail","sn","givenName"
		$arrResults.Add($(
			New-Object PSObject -Property @{
				"AccountIsEnabled" = $objADUser.enabled;
				"Surname" = $objADUser.sn;
				"GivenName" = $objADUser.givenName;
				"DisplayName" = $objADUser.displayName;
				"SamAccountName" = $objADUser.samAccountName;
				"EmailAddress" = $objMailbox.PrimarySMTPAddress;
				"Location" = $objADUser.l;
				"Department" = $objADUser.department;
				"Title" = $objADUser.title;
				"OrganizationalUnit" = $objMailbox.OrganizationalUnit;
				"ItemCount" = $objSearchResults.SearchResult.ItemCount;
				"ItemSize" = $objSearchResults.SearchResult.Size;
			} | Select "AccountIsEnabled","Surname","GivenName","DisplayName","SamAccountName","EmailAddress","Location","Department","Title","OrganizationalUnit","ItemCount","ItemSize"
		)) > $null
	}
}
Write-Progress -Activity "Searching Mailboxes" -Id 1 -Completed

Remove-PSSession $objPSSession
Remove-Variable objPSSession

$arrResults