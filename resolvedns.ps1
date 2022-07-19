

#Get-Content "D:\Scripts\ResolveDNS\records.txt" | Resolve-DNSName | Select Name,IPAddress | Out-File "D:\Scripts\ResolveDNS\result3.txt"

#Get-Content "D:\Scripts\ResolveDNS\records2.txt" | Resolve-DNSName | Out-File "D:\Scripts\ResolveDNS\result2.txt"

Get-Content D:\Scripts\ResolveDNS\records2.txt | Test-Connection
