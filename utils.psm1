<#
.synopsis
    * Adds an additional hashtable to a first, overwriting previous keys if they exist. Respects ordering
.description
    * Adds an additional hashtable to a first, overwriting previous keys if they exist. Respects ordering
.parameter Original
    * The first hashtable. Keys in this will be overwritten by the second if they exist there
.parameter Additional
    * The hashtable to add
.notes
    * Original Author: iRon
    * Email: 
    * From: https://stackoverflow.com/questions/8800375/merging-hashtables-in-powershell-how
    * Updated:
      * By GSGBen
    * * Allowed ordered hash tables
    * Notes: not sure why parameters don't work. Pass by ref?
#>
Function Merge-HashTable
{
    $Output = [ordered]@{}
    ForEach ($Hashtable in ($Input + $Args)) {
            ForEach ($Key in $Hashtable.Keys) {$Output.$Key = $Hashtable.$Key}
    }
    $Output
}