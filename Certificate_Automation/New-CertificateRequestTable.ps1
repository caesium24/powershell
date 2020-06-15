<#
 .NOTES
    Version  : 1.0
    Author   : Eshton Brogan & Sid Johnston
    Created  : 15 June 2020
 
 .SYNOPSIS
  Generates a spreadsheet of certificate requests and their status which can be submitted to Trusted Agents for certificate request approval.

 .DESCRIPTION
  Utilizing user-provided information and a links CSV, this function outputs a spreadsheet which contains all data relevant to Trusted
  Agents in a user-friendly format. 
  
 .PARAMETER CSRFilePath
  File path containing source CSR(s).

 .PARAMETER LinkFilePath
  File path of links.csv which contains the status URL(s) for submitted certificate request(s).
  
 .PARAMETER CA
  Name of the Certificate Authority which will issue the certificate(s).
  
 .PARAMETER Region
  Name of the city or locality in which the certificate host(s) will reside.
  
 .PARAMETER Base
  Name of the base or facility in which the certificate host(s) will reside.

 .PARAMETER NetworkSysAdmin
  Name of the person who requested the certificate(s).

 .PARAMETER SystemOwner
  Name of the project or program in which the certificate host(s) will reside.

 .PARAMETER SecurityLevel
  Classification level of the certificate host(s).

 .PARAMETER ResolveDNSNames
  Boolean switch to attempt to resolve DNS A records for source hostname(s).

 .EXAMPLE
  New-CertificateRequestTable -CSRFilePath "C:\temp\CSR" -LinkFilePath C:\temp\links.csv -CA "CA-1" -Region "Panama City" -Base "NSWC" -NetworkSysAdmin "John Smith" -SystemOwner "Site01, Panama City, FL" -SecurityLevel "UNCLASSIFIED" -ResolveDNSNames:$false 
#>
function New-CertificateRequestTable {
    [CmdletBinding()]
    param (
        [Alias("CSRPath")]
        [Parameter(Mandatory=$true)]
        [String]
        $CSRFilePath,
        [Alias("LinkPath")]
        [Parameter(Mandatory=$true)]
        [String]
        $LinkFilePath,
        [Parameter(Mandatory=$true)]
        [String]
        $CA,
        [Parameter(Mandatory=$true)]
        [String]
        $Region,
        [Parameter(Mandatory=$true)]
        [String]
        $Base,
        [Parameter(Mandatory=$true)]
        [String]
        $NetworkSysAdmin,
        [Parameter(Mandatory=$true)]
        [String]
        $SystemOwner,
        [Parameter(Mandatory=$true)]
        [String]
        $SecurityLevel,
        [Parameter(Mandatory=$false)]
        [bool]
        $ResolveDNSNames = $false
    )

    function Get-IndicesOf ($Array, $Value) {
        $i = 0
        foreach ($element in $Array) {
            if ($element -eq $Value) {$i}
            ++$i
        }
    }

    if ((Test-Path -Path $CSRFilePath -PathType Container) -eq $true){
        if ((Test-Path -Path $LinkFilePath -PathType Leaf) -eq $true){
            $requestscsv = @()
            $link_csv = Import-Csv -Path $LinkFilePath
            $requests = Get-ChildItem -Path $CSRFilePath | Where-Object {$_.Name.EndsWith('.csr') -eq $true}
            $requests = $requests.Name
            foreach ($request in $requests) {
                $dns_name = $request.TrimEnd(".csr")
                    if ($ResolveDNSNames -eq $true){
                        $ip = Test-NetConnection -ComputerName $request
                    }
                foreach ($el in $link_csv){
                    if($el.Hostname -eq $dns_name) {
                        $linkRequest = $el.RequestID
                        $linkID = ($linkRequest -split "=")[1]
                    }
                }

                $obj = @(
                
                    [PSCustomObject]@{
                        CA = $CA
                        RequestID = $linkID
                        RequestLink = $linkRequest
                        Hostname = $dns_name
                        IPaddress = $(if ($ResolveDNSNames -eq $true){$ip.RemoteAddress.IPAddressToString} else {''})
                        Region = $Region
                        Base = $Base
                        NetworkSysAdmin = $NetworkSysAdmin
                        SystemOwner = $SystemOwner
                        SecurityLevel = $SecurityLevel
                        AppOnServer = ""
                        Notes = ""
                    }
                )

            $requestscsv += $obj
            }

            $requestscsv | Export-Csv -Path "$CSRFilePath\CertificateRequestTable.csv"
        }
        else {
         throw "Error: $LinkFilePath file not found"
        }
    }
    else {
        throw "Error: $CSRFilePath path not found"
    }
}
