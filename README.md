# XLSX-to-Org-Table-in-Emacs

Barebone example to read within an Excel file the 1st tab/sheet, convert into an org Table for Emacs. 

The setup is basic and relies on:
1) Powershell --> Tested only on W10 but should work on the other new platform MacOs and Unix Microsoft is targeting as long as the key dependency 
Import-Excel library is working
2) A module Import-Excel which can be installed 
3) A powershell function to be found in the XLTable2Org.ps1 file
4) A few lines of Elisp to be added to your init.el .emacs or equivalent in order to fire the function.



