<#
 .NOTES
    Version : 1.0
    Author  : Eshton Brogan & Sid Johnson
    Created : 15 June 2020
    
 .SYNOPSIS
  Submits certificate requests to a DISA Certificate Authority website and returns a CSV containing links to the certificate request status pages.
  
 .DESCRIPTION
  This function submits certificate requests to a DISA Certificate Authority website. Given a source for CSR files, each CSR is submitted
  in turn to the user-supplied CA website along with the name, email address, and phone number of the requestor. As each request is
  submitted, a CSV containing links to the status pages of each certificate request is populated for later use and output to the path
  specified in the function.
  
 .PARAMETER CSRFolderPath
  Folder path for the new CSR files to be generated.
  
 .PARAMETER CA
  Base URL for the DISA Certificate Authority website.
  
 .PARAMETER SavePath
  Folder path for the CSV containing links to the certificate request status pages.
  
 .PARAMETER RequestorName
  Name of the person making the certificate request(s).
  
 .PARAMETER RequestorEmail 
  Email address of the person making the certificate request(s).
  
 .PARAMETER RequestorPhone
  Phone number of the person making the certificate request(s).
  
  NOTE: This parameter will only accept string input and must not contain dashes ( - ). See example for further detail. 
 
 .EXAMPLE
  Request-DisaCertificate -CSRFolderPath "C:\temp\CSR" -CA https://casite.csd.disa.mil -SavePath C:\temp -RequestorName "John Smith" -RequestorEmail "john.smith@navy.mil" -RequestorPhone "1235557890"
#>
function Request-DisaCertificate {
    [CmdletBinding()]
    param (
        # Folder Path for Certificate Requests
        [Parameter(Mandatory=$true)]
        [String]
        $CSRFolderPath,
        # URL to CA
        [Parameter(Mandatory=$true)]
        [String]
        $CA,
        # Save Path for Link File
        [Parameter(Mandatory=$true)]
        [String]
        $SavePath,
        # Requestor's name
        [Parameter(Mandatory=$true)]
        [String]
        $RequestorName,
        # Requestor's email
        [Parameter(Mandatory=$true)]
        [String]
        $RequestorEmail,
        # Requestor's phone number
        [Parameter(Mandatory=$true)]
        [String]
        $RequestorPhone
    )
    
    begin{
        $links = @()
        $headers = @{
            'Accept' = 'text/html, application/xhtml+xml, image/jxr, */*'
            'Referer' = "$($CA)/ca/ee/ca/profileSelect?profileId=cspMultiSANCert"
        }

        if ((Test-Path -Path $CSRFolderPath -PathType Container) -eq $true){
            $requests = Get-ChildItem -Path $CSRFolderPath | Where-Object {$_.Name.EndsWith('.csr') -eq $true}
        }
        else {
            throw "Error: Folder path not found"
        }
    }

    process {
    [Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
        foreach ($request in $requests){
            $csr = Get-Content -Path $request.VersionInfo.Filename -Raw
            $dns_name = $request.Name.TrimEnd('.csr')
            $body = @{
                'cert_request_type' = 'pkcs10'
                'cert_request' = $csr
                'gnameValue1' = $dns_name
                'requestor_name' = $RequestorName
                'requestor_email' = $RequestorEmail
                'requestor_phone' = $RequestorPhone
                'profileId' = 'cspMultiSANCert'
                'renewal' = 'false'
                'xmlOutput' = 'false'
            }

            try {
                Invoke-WebRequest -Uri "$($CA)/ca/ee/ca/profileSubmit" -Method Post -OutVariable webresponse -Headers $headers -Body $body -ErrorAction Stop
                $obj = @(
                    [PSCustomObject]@{
                        Hostname = $dns_name
                        RequestID = "$($CA)/ca/ee/ca/$($webresponse.links.href[2])"
                    }
                    )
                $links += $obj
            }
            catch [System.Management.Automation.RuntimeException]{
                Write-Warning -Message "Cannot resolve URI"
            }
        }
    }
    end {
        if ((Test-Path -Path "$SavePath\links.csv" -PathType Leaf) -eq $true){
            $links | Export-Csv -Path "$SavePath\links.csv -Append"
        }
        else {
            $links | Export-Csv -Path "$SavePath\links.csv"
        }
    }
}
