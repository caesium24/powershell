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
        $webreq.ParsedHtml.IHTMLDocument3_documentelement.outerText | Out-file -FilePath $SavePath
    }
}
