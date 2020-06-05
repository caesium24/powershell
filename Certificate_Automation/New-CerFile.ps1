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
