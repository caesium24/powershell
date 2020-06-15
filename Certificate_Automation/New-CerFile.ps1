<#
 .NOTES
    Version  : 1.0
    Author   : Eshton Brogan & Sid Johnston
    Created  : 15 June 2020
 
 .SYNOPSIS
  Parses raw HTML files for certificate hashes and outputs CER files.

 .DESCRIPTION
  Utilizing user-provided file paths, this function parses the certificate hashes from raw HTML files and
  returns CER files for use as host certificates. 
  
 .PARAMETER TextFolderPath
  Folder path of raw HTML text files which contain certificate hashes.

 .PARAMETER CerFolderPath
  Folder path to output parsed CER files.
  
 .EXAMPLE
  New-CerFile -TextFolderPath C:\temp\full -CerFolderPath C:\temp\CER
#>
function New-CerFile {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]
        $TextFolderPath,
        [Parameter(Mandatory=$true)]
        [String]
        $CerFolderPath
    )
    $certFiles = Get-ChildItem -Path $TextFolderPath
    foreach ($file in $certFiles) {
        $name = ($file.name).trimend(".txt")
        $txt = Get-Content -Path "$($TextFolderPath)\$($file)"
        $txt = $txt[118..146]
        $txt | Out-File -FilePath "$($CerFolderPath)\$($name).cer" -Encoding "ASCII"
    }
}
