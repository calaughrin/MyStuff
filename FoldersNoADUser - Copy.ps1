import-module ActiveDirectory

$FolderPath = "EnterLocationHere"
$UnusedFolderLogPath = "DesiredLocationHere\FoldersNoADUser.txt"
$ErrorLog = "DesiredLocationHere\ErrorLogFoldersNoADUser.txt"


#Gets Each Folder In RootDirectory, which is each users home directory that is names the same name has there SamAccountName
$folders = Get-ChildItem $FolderPath 

#Loops Through Each Home Directory Folder
ForEach ($folder in $folders) 
{
    #Searches For User
    $User = Get-ADUser -LDAPFilter "(sAMAccountName=$folder)"

    #Checks If Users Exists
    If ($User -eq $Null)
    {
    $HomeDir = $FolderPath+"\"+$folder
    Then
    $LogEntryDate = Get-Date –Format G
                
                #Log Source Path Was Not Removed
                Add-Content -Path $UnusedFolderLogPath -Value "******************************************************************************************************************"
                Add-Content -Path $UnusedFolderLogPath -Value "Log Entry Date: $LogEntryDate"
                Add-Content -Path $UnusedFolderLogPath -Value "\\\///User: $folder\\\///"
                Add-Content -Path $UnusedFolderLogPath -Value "This Folder does not have a matching user."
    }

}