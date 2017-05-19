ipmo ActiveDirectory
ipmo C:\Scripts\get-password-anders.ps1

Import-Csv C:\Scripts\UserCreationList.csv | foreach-object { 
$userprinicpalname = $_.SamAccountName + “@domain.edu”
$password= Get-Password 
New-ADUser -SamAccountName $_.SamAccountName -UserPrincipalName $userprinicpalname -Name $_.name -DisplayName $_.name -GivenName $_.cn -SurName $_.sn -Department $_.Department -Path “OU Path” -AccountPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Enabled $True -AccountExpirationDate 12/02/2016 -PassThru }

Get-ADUser -Filter * -SearchBase “OU Path” -Properties displayName | select displayName | ForEach-Object {Get-Password} | Out-File -filepath C:\Scripts\CreatedUserList.txt