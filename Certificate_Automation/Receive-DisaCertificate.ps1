<#
 .NOTES
    Version  : 1.0
    Author   : Eshton Brogan & Sid Johnston
    Created  : 15 June 2020
 
 .SYNOPSIS
  Pulls raw HTML for completed certificate requests from a specified Certificate Authority website.

 .DESCRIPTION
  Utilizing user-provided information and a Certificate Request CSV, this function outputs raw text files which contain
  certificate hash values. 
  
 .PARAMETER CSVPath
  File path of CSV which contains the status URL(s) for submitted certificate request(s).

 .PARAMETER CA
  Base URL for the DISA Certificate Authority website.
  
 .PARAMETER SavePath
  File path where raw HTML files will be saved. 

 .EXAMPLE
  Receive-DisaCertificate -CSVPath C:\temp\certificate_request.csv -CA "http://ca-1.csd.disa.mil" -SavePath C:\temp\full 
#>
function Receive-DisaCertificate {
    [CmdletBinding()]
    param (
        # File Path for Certificate Request CSV
        [Parameter(Mandatory=$true)]
        [String]
        $CSVPath,
        # URL to CA
        [Parameter(Mandatory=$true)]
        [String]
        $CA,
        # Save Path for retrieved certificate
        [Parameter(Mandatory=$true)]
        [String]
        $SavePath
    )
  [Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
  $requests =Import-Csv -Path $CSVPath

    foreach ($cert in $requests){
        $reqID = $cert.requestID
        $name = $cert.hostname

        $webreq = Invoke-WebRequest -Uri "$($CA)/ca/ee/ca/checkRequest?requestId=$($reqID)"
        $link = $webreq.links.href[1]
        $webreq = Invoke-WebRequest -Uri "$($CA)/ca/ee/ca/$($link)"
        $webreq.ParsedHtml.IHTMLDocument3_documentelement.outerText | Out-file -FilePath "$SavePath\$($name).txt"
    }
}
