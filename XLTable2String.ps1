# ------------------------------------------------------------------------------
#                     Time-stamp: "2021-01-18 11:26:27 jpdur"
# ------------------------------------------------------------------------------
# Actual Location of Microsoft.Powershell_profile.ps1 is c:\Users\XXXXX\Documents\WindowsPowerShell
# where XXXX is the name of the User (as per Windows) 

# Read an XL file, extract a table and presents it into an org compatible table format
function XLTable2String($Dest)
{
    # # Extract the Table into a csv file for Debug only
    # Import-Excel $Dest | Export-Csv Test2.csv -Delimiter '|' -NoTypeInformation 

    # Extract the Table into a String with | delimiters between the fields
    $TextTable = Import-Excel $Dest | ConvertTo-Csv -Delimiter '|' -NoTypeInformation 

    # It comes as follows "Lisa"|"Mum"|"62"
    $TextTable = $TextTable.replace('"|"','|')

    # Replace the " at the beginning and end of file 
    # !!!! .replace does not process regexp but -replace does !!!!!
    $TextTable = $TextTable -replace('^"(.*?)"$','|$1|')

    return $TextTable 
}
