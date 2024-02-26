# Exchange Retention What/If Analysis
PowerShell script that performs a what/if analysis on mailboxes if a defined retention policy was applied.  The script produces a report of each mailbox with a total item size and count of what would be removed if the given retention policy was applied to the mailbox, as well as information about the associated Active Directory user account, like Department, Title, Location, etc.

## Usage
```console
.\ExchangeRetentionWhatIfAnalysis.ps1 -?
```
Requires ActiveDirectory PowerShell module and Exchange Web Services.

## Examples
```console
.\ExchangeRetentionWhatIfAnalysis -SearchDate $((Get-Date).AddYears(-2)) -ExchangeHost exch01.company.com -EWSURI https://ews.company.com/EWS/Exchange.asmx
```
Runs the report against all mailboxes and calculates item size/count based on a 2 year retention policy.

```console
.\ExchangeRetentionWhatIfAnalysis -SearchDate $((Get-Date).AddDays(-90)) -MailboxName user1,user2 -ExchangeHost exch01.company.com -EWSURI https://ews.company.com/EWS/Exchange.asmx
```
Runs the report against the "user1" and "user2" mailboxes and calculates item size/count based on a 90 day retention policy.
