# Wiki

Here: https://github.com/enderphan94/Learn-Power-Shell/wiki

# Problems reported

Here: https://github.com/enderphan94/Learn-Power-Shell/wiki#anatomy-ldap-attributes

# Forensic: Active Directory Account Types

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

- Powershell Version 4 as a minimum 
- Import-Module ActiveDirectory [How to import AD Module](https://github.com/enderphan94/Learn-Power-Shell/wiki#how-to-install-ad-cmdlets)
- < Windows 7/ Windows Server 2012..
- Run as regular user

```
The tool has been tested in Windows Server 2012
```

### Installing

A step by step series of examples that tell you have to get a development env running

1. Clone it to your directory:

    `git clone https://github.com/enderphan94/ad-Tracking/`
    
2. Run as a regular user:

    `./ad-seeking.ps1`
    
3. Follow the prompts given from the tool

## Deployment

In order to deploy this on a live system, you don't need an Admin account. You can deploy it from the trusted Domain Controller where it's able to query the data from others needed DCs.

## Built With

1. Powershell 
2. .Net 
3. HTML5/CSS

## Versioning

Last version updated! - Version 1.0

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

Given out HTML reports with its illustration displayed in Bar charts and Column graphs.

## Authors

* **Ender Loc Phan** - *Initial work* - [GitRespo](https://github.com/enderphan94)

See also the list of [CCI-WebPage](http://enderphan.2wy.info) projects I have done 

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Hat tip to anyone who's code was used
* Inspiration
* etc
