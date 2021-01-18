# XLSX-to-Org-Table-in-Emacs

Barebone example to read within an Excel file the 1st tab/sheet, convert into an org Table for Emacs. 

The setup is basic and relies on:

1) Powershell --> Tested only on W10 but should work on the other new platform MacOs and Unix Microsoft is targeting as long as the key dependency 
Import-Excel library is working

2) A module Import-Excel which can be installed as per the indication given at https://www.powershellgallery.com/packages/ImportExcel/7.0.1

3) A powershell function to be found in the XLTable2Org.ps1 file
It is easier to add the Powershell function to the profile so that way the PS script is available independently of where the scruipt is being called. It appears as a new function of Powershell. 

In order to add the function to the Powershell profile the simplest option is to manually add the function into the Microsoft.Powershell_profile.ps1 file which is usually in the C:/Users/XXXX/Documents directory. 

4) A few lines of Elisp to be added to your init.el .emacs or equivalent in order to fire the function.

There is no error trapping and it can be said that this is more a Proof of Concept than anything else but it actually works....


