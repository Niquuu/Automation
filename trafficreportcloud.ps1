# Connect cloud service! Import modules if you dont have those.

<# 

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned 

Import-Module ExchangeOnlineManagement 


#>

Connect-ExchangeOnline -UserPrincipalName "User" -DelegatedOrganization "Organization"

# Collects Mail traffic report from domain and makes CSV file 

try {

$end = Get-Date -Format MM/dd/yyyy
$start = $(Get-Date).addMonths(-1).ToString("MM/dd/yyyy")


Get-MailTrafficTopReport -Domain "Domain"  -StartDate $start -EndDate $end -PageSize 5000 `
| Select-Object name, MessageCount, Direction | Where name -EQ "email" `
| Export-csv -NoTypeInformation ".\File.csv"


Get-MailTrafficTopReport -Domain "Domain"  -StartDate $start -EndDate $end -PageSize 5000 `
| Select-Object name, MessageCount, Direction, (Date).toString("dd/MM/yyyy") | Where name -EQ "email" `
| Export-csv -NoTypeInformation ".\File.csv"


Get-MailTrafficTopReport -Domain "Domain" -StartDate $start -EndDate $end -PageSize 5000 `
| Select-Object name, MessageCount, Direction, (Date).toString("dd/MM/yyyy") | Where name -EQ "email" `
| Export-csv -NoTypeInformation ".\File.csv"


Get-MailTrafficTopReport -Domain "Domain"  -StartDate $start -EndDate $end -PageSize 5000 `
| Select-Object name, MessageCount, Direction, (Date).toString("dd/MM/yyyy") | Where name -EQ "email" `
| Export-csv -NoTypeInformation ".\File.csv"


Get-MailTrafficTopReport -Domain "Domain"  -StartDate $start -EndDate $end -PageSize 5000 `
| Select-Object name, MessageCount, Direction, (Date).toString("dd/MM/yyyy") | Where name -EQ "email" `
| Export-csv -NoTypeInformation ".\File.csv" 


Get-MailTrafficTopReport -Domain "Domain"  -StartDate $start -EndDate $end -PageSize 5000 `
| Select-Object name, MessageCount, Direction, (Date).toString("dd/MM/yyyy") | Where name -EQ "email" `
| Export-csv -NoTypeInformation ".\File.csv"


# Calculate all CSV files Indound and Outbound count together and make own file.

$data = Import-Csv ".\File.csv";

$data | group Name | select Name, @{n="Inbound";e={(($_.Group | where Direction -eq "Inbound").MessageCount | Measure-Object -Sum).Sum}}, @{n="Outbound";e={(($_.Group | where Direction -eq "Outbound").MessageCount | Measure-Object -Sum).Sum}} | Out-File .\test.txt

$data = Import-Csv ".\File.csv";

$data | group Name | select Name, @{n="Inbound";e={(($_.Group | where Direction -eq "Inbound").MessageCount | Measure-Object -Sum).Sum}}, @{n="Outbound";e={(($_.Group | where Direction -eq "Outbound").MessageCount | Measure-Object -Sum).Sum}} | Out-File .\test1.txt

$data = Import-Csv ".\File.csv";

$data | group Name | select Name, @{n="Inbound";e={(($_.Group | where Direction -eq "Inbound").MessageCount | Measure-Object -Sum).Sum}}, @{n="Outbound";e={(($_.Group | where Direction -eq "Outbound").MessageCount | Measure-Object -Sum).Sum}} | Out-File .\test2.txt

$data = Import-Csv ".\File.csv";

$data | group Name | select Name, @{n="Inbound";e={(($_.Group | where Direction -eq "Inbound").MessageCount | Measure-Object -Sum).Sum}}, @{n="Outbound";e={(($_.Group | where Direction -eq "Outbound").MessageCount | Measure-Object -Sum).Sum}} | Out-File .\test3.txt

$data = Import-Csv ".\File.csv";

$data | group Name | select Name, @{n="Inbound";e={(($_.Group | where Direction -eq "Inbound").MessageCount | Measure-Object -Sum).Sum}}, @{n="Outbound";e={(($_.Group | where Direction -eq "Outbound").MessageCount | Measure-Object -Sum).Sum}} | Out-File .\test4.txt

$data = Import-Csv ".\File.csv";

$data | group Name | select Name, @{n="Inbound";e={(($_.Group | where Direction -eq "Inbound").MessageCount | Measure-Object -Sum).Sum}}, @{n="Outbound";e={(($_.Group | where Direction -eq "Outbound").MessageCount | Measure-Object -Sum).Sum}} | Out-File .\test5.txt


# Collect all data from new files and put them together

Get-Content Test*.txt | Set-Content "MonthTotal $(get-date -f dd/MM/yyyy).txt"

} catch {

Write-Warning "Someting is wrong. Check code!"

}
