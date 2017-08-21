# ad-Seeking

Dev by Ender Loc Phan

# Requirements:

- Powershell > Version 4
- Import-Module ActiveDirectory
    
# DESCRIPTION

    Get Service, Admin, User accounts which never logon and have not logged on in 105 days with the following attributes:

    "distinguishedName",
    "sAMAccountName",
    "mail",
    "lastLogonTimeStamp",
    "pwdLastSet",
    "badpwdcount",
    "accountExpires",
    "userAccountControl",
    "modifyTimeStamp",
    "lockoutTime"
    "badPasswordTime",
    "maxPwdAge ",
    "Description"
    
# NOTES
    Version 1.0
# LINK
    https://github.com/enderphan94/ad-Seeking/blob/master/ad-seeking.ps1
    
# Usage
 
- Following the instruction given by the tool
- Run quickCheck.ps1 to compare the ouput with amout of data
