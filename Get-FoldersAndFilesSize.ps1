<#
    .SYNOPSIS
    Gets files and folders size.

    .DESCRIPTION
    Gets files and folders size.
    Also queried for files and folders size and also for folders and files count (for folders only). Final step is export results to CSV file.

    .PARAMETER FolderPath
    Specifies the root folder that will be recursively queried for files and folders size and also for folders and files count (for folders only). 

    .PARAMETER ExportCsvPath
    Specifies the path for output CSV file.

    .INPUTS
    None. You cannot pipe objects to Add-Extension.

    .OUTPUTS
    [System.IO.File]. Get-FoldersAndFilesSize.ps1 returns a CSV file with the result of execution.
    Example:
    Root Folder;;Folder;;Size (bytes);;Total Objects;;Folders Count;;Files Count;;Type;;Notes
    C:\Profiles\;;LargeFiles9_2022-02-07_00-21-43_files;;18753;;4;;0;;4;;Directory;;
    C:\Profiles\;;LeastRecentlyAccessed11_2022-02-11_03-50-54_files;;7649;;2;;0;;2;;Directory;;
    C:\Profiles\;;New folder;;0;;1;;1;;0;;Directory;;
    C:\Profiles\;;QuotaUsage2_2022-02-03_21-53-44.html;;4899;;0;;0;;0;;File;;

    .EXAMPLE
    PS> .\Get-FoldersAndFilesSize.ps1 -FolderPath C:\Profiles\ -ExportCsvPath C:\tmp\file.csv

    .LINK
    Get-FoldersAndFilesSize.ps1
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$FolderPath,

    [Parameter(Mandatory=$true)]
    [string]$ExportCsvPath
)

New-Item -Path $ExportCsvPath -Force -ItemType File -Value '' | Out-Null
Add-Content -Path $ExportCsvPath -Value 'Root Folder;;Folder;;Size (bytes);;Total Objects;;Folders Count;;Files Count;;Type;;Notes'

$items = Get-ChildItem -Path $FolderPath

foreach($item in $items){
    $path = $FolderPath + $item

    $size = 0
    $foldersCount = 0
    $filesCount = 0
    $totalObjectCount = 0
    $notes = ''

    try
    {
       if ((Get-Item $path) -is [System.IO.DirectoryInfo]){
            $foldersCount = (Get-ChildItem -Recurse -Path $path -Directory -ErrorAction SilentlyContinue | Measure-Object).Count 
            $filesCount = (Get-ChildItem -Recurse -Path $path -File -ErrorAction SilentlyContinue | Measure-Object).Count
            $totalObjectCount = $foldersCount + $filesCount

            if ($filesCount -gt 0){
                $size = (Get-ChildItem -Path $path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum   
            }

            $valueToAdd = $FolderPath + ";;"  + $item + ";;" + $size + ";;" + $totalObjectCount + ";;" + $foldersCount + ";;" + $filesCount + ";;" + "Directory" + ";;" + $notes
            Add-Content -Path $ExportCsvPath -Value $valueToAdd 
       } else {
            $size = (Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum
            $valueToAdd = $FolderPath + ";;"  + $item + ";;" + $size + ";;" + $totalObjectCount + ";;" + $foldersCount + ";;" + $filesCount + ";;" + "File" + ";;" + $notes
            Add-Content -Path $ExportCsvPath -Value $valueToAdd
       }
         
    }

    catch
    {
        $notes = "Failed to process item: " + $item
    }   

}