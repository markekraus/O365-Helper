# O365-Helper
Office 365 PowerShell Helper Script

This PowerShell Script is designed to provide several helper functions for managing Office 365 Exhcnage Online, SharePoint Online and OneDrive for buesiness. The goal is to create commandlets that simplify common tasks done in the Office 365 environment. 

In order to take full advantage of this script you will need to install the Following:

Microsoft Online Services Sign-In Assistant:  
* http://www.microsoft.com/en-us/download/details.aspx?id=41950

Azure AD Module for Windows PowerShell:  
* 64-Bit: http://go.microsoft.com/fwlink/p/?linkid=236297  
* 32-Bit: http://go.microsoft.com/fwlink/p/?linkid=236298  
    
SharePoint Online Management PowerShell Module:  
* https://www.microsoft.com/en-us/download/details.aspx?id=35588  
 
# Installation and Use

* Download `O365-Helper.ps1`
* Save it to a folder of your chosing (e.g. `C:\O365-Helper\`)
* Modify the Variables in the Configuration section of `O365-Helper.ps1`
* Source the file in your shell, scripts, or profile:
```PowerShell
. "C:\O365-Helper\O365-Helper.ps1"
```
* Connect to Office 365:  
```PowerShell
Connect-O365
```  
* Use any of the commandlets
```PowerShell
Get-O365DistributionGroupMembership -Identity "bob.testerton@contoso.com"
```

