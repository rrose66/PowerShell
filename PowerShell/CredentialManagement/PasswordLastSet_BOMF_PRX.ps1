Get-ADUser BOMF_PRX  –Properties * | select passwordlastset | Out-File -Encoding ASCII  -FilePath C:\PROJECTS\DATA\FEDEBOM\CREDENTIAL_MANAGEMENT\CREDENTIAL_PASSWORD_MANAGEMENT\AD_SEARCH_RESULTS\PasswordLastSet_BOMF_PRX.txt