mode 300
Get-ChildItem | Select-Object -Property CreationTime, Name | Out-File -FilePath output.txt
exit