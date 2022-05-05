# Get-Marval-FileReport.ps1
<#
Author: Björn Gustavsson, Westcode.se
Right click .ps1 file and Open with PowerShell
Script will look up regvalue for Marval Installation 
and generate a file report with:

Filname, Hash, FileSize
 
#>

#Get Marval Regkey
$marval = (Get-ItemProperty -Path Registry::"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Marval Software\MSM" 2>&1 $null)
if (!$marval.path){
write-warning "Regkey HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Marval Software\MSM was not found"
break
}
#Define Marval install path
Write-Host "Found Path:" $marval.Path
$version = $marval.Version
Write-Host "Collecting files"
#Empty Array 
$array = @()
# Find files in sub-folders
$files = Get-ChildItem $marval.Path -Attributes !Directory -Recurse -Exclude *.eml,*.log
 
# Calculate size in MB for files
$size = $Null
$files | ForEach-Object -Process {
Write-Progress -Activity "Processing Marval Files" -CurrentOperation $_.FullName -PercentComplete (($counter / $files.Count) * 100)

$file = $_.FullName
$size = $_.Length
$counter++
 
$md5 = (Get-FileHash -Algorithm md5 $file).hash
$sizeinmb = [math]::Round(($size / 1mb), 1)
$sizeinkb = [math]::Round(($size / 1kb), 1)
# Add pscustomobjects to array
$array += [pscustomobject]@{
File = $file
Filehash = $md5
'Size(MB)' = $sizeinmb
'Size(KB)' = $sizeinkb
}

}

# Generate Report Results in your Current Folder
$array|Export-Csv -Path $env:COMPUTERNAME-Marval_$version-file_report.csv -NoTypeInformation
