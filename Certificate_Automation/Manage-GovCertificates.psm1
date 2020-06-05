########################################################################################################################################################################
<#

Creates new SSL Certificate Requests and Keys using either vCenter, CSV, or user input to populate hostnames

    Version : 1.6
    Author  : Eshton Brogan & Sid Johnson
    Created : 09 October 2019
    
  .Synopsis
  Placeholder

 .Description
  Placeholder

 .Parameter CSRPath
  File path for new CSR files to be generated.

 .Parameter KeyPath
  File path for new KEY files to be generated.

 .Parameter Source
  Parameter to choose between vCenter, Hostname, and File sources for certificate hostnames.
  
 .Parameter vCenterServer
  If vCenter is selected as a source, this dynamic parameter specifies the vCenter server to source the certificate hostnames.
  
 .Parameter Hostname 
  If Hostname is selected as a source, this dynamic parameter specifies the hostname of the certificate being generated.
  
 .Parameter FilePath
  If File is selected as a source, this dynamic parameter specifies the CSV file which contains the certificate hostnames.
  
 .Parameter Credential
  Credentials for vCenter if it is selected as a source.
 
 .Example
  Placeholder

 .Example
  Placeholder

 .Example
  Placeholder
#>
########################################################################################################################################################################
function New-CertificateRequest {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String[]]
        $CSRPath,
        [Parameter(Mandatory=$false)]
        [String[]]
        $KeyPath,
        [Parameter(Mandatory=$true)]
        [ValidateSet('vCenter', 'Hostname', 'File')]
        [String[]]
        $Source,
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]
        $Credential
    )
        DynamicParam {
            switch ($Source){
                "vCenter" {$paramName = "vCenterServer"}
                "Hostname" {$paramName = "Hostname"}
                "File" {$paramName = "FilePath"}
            }
    
            $attributes = New-Object -TypeName System.Management.Automation.ParameterAttribute
            $attributes.ParameterSetName = "dynSet"
            $attributes.Mandatory = $true
            $attributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)
    
            $srcParam = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter($paramName, [String], $attributeCollection)
    
            $paramDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add($paramName, $srcParam)
            return $paramDictionary
        }
        begin {
            $openssl = 'C:\Program Files\OpenSSL-Win64\bin\openssl.exe'
            #$openssl = where.exe /R 'C:\Program Files' openssl.exe | Select-Object -First 1
                if ($null -eq $openssl){
                    Write-Warning -Message "Openssl not found in C:\Program Files directory. Searching again in C:\"
                    #$openssl = where.exe /R 'C:\' openssl.exe | Select-Object -First 1
                    $openssl = 'C:\Program Files\OpenSSL-Win64\bin\openssl.exe'
                        if($null -eq $openssl){throw "Error: Openssl.exe not found on localhost"}
                }
                $openssl = $openssl.TrimEnd('openssl.exe')
                $noName = @()
        }
        process {
            if ($Source -eq "vCenter") {
                $mods = Get-Module -ListAvailable | Where-Object {$_.Name -like "VMware*"}
                $installed = Get-Module -All | Where-Object {$_.Name -like "*VMware*"}
                if ($null -ne $installed) {
                    Write-Host -ForegroundColor Green "Vmware Modules already available."
                }
                elseif ($null -eq $mods){
                   Import-Module -Name $mods -ErrorAction Stop
                }
    
                try {
                   Connect-ViServer -Server $PSBoundParameters.vCenterServer -Credential $Credential -ErrorAction Stop
                }
                catch [VMware.VimAutomation.Sdk.Types.V1.ErrorHandling.VimException.ViServerConnectionException]{
                   Write-Warning -Message "Cannnot find vCenter/ESX/ESXi server. Double check IP address or if localhost can resolve the server hostname"
                }
                catch [VMware.VimAutomation.ViCroe.Types.V1.ErrorHandling.InvalidLogin]{
                   Write-Warning -Message "Incorrect Username/Password"
                }
    
                $vmList = Get-VM | Get-VMGuest
                Set-Location $openssl
                foreach ($vm in $vmList) {
                    if ($null -ne "$vm.HostName") {
                        $vmname =$vm.HostName.ToLower()
                        .\openssl.exe req -nodes -newkey -rsa:2048 -sha256 -nodes -keyout "$($KeyPath)\$($vmname).key" -out "$($CSRPath)\$($vmname).csr" -subj /CN=$($vmname)/OU=NSS/O=PKI/ST=DOD/L=U.S. Government/C=US
                    }
                    else {
                        $noName += $vm.VmName
                    }
                }
                $noName | Out-File -FilePath "$CSRPath\MissingHostNamesList.csv"
    
                Write-Warning -Message "The Following Virtual Machines do not have a listed Hostname/DNS Name in vCenter. Please check that they have VMware tools installed, or correct the MissingHostNamesList.csv file in $CSRPath and rerun this Function using the -Source File and -FilePath parameters."
            }
            elseif ($Source -eq "File") {
                Set-Location $openssl
                $FilePath = $PSBoundParameters.FilePath
                $vmList = Import-Csv -Path "$FilePath"
                foreach ($vm in $vmList) {
                    .\openssl.exe req -nodes -newkey -rsa:2048 -sha256 -nodes -keyout "$($KeyPath)\$($vm).key" -out "$($CSRPath)\$($vm).csr" -subj /CN=$($vm)/OU=NSS/O=PKI/ST=DOD/L=U.S. Government/C=US            
                }
            }
            elseif ($Source -eq "HostName") {
                Set-Location $openssl
                $HostName = $PSBoundParameters.HostName
                .\openssl.exe req -nodes -newkey -rsa:2048 -sha256 -nodes -keyout "$($KeyPath)\$($HostName).key" -out "$($CSRPath)\$($HostName).csr" -subj /CN=$($HostName)/OU=NSS/O=PKI/ST=DOD/L=U.S. Government/C=US     
            }
            else {
                Write-Error "Could not create certificate requests based on the information provided."
            }
        }
    }
##################################################################################################################################

function Receive-DisaCertificate {
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
        # Requester's name
        [Parameter(Mandatory=$true)]
        [String]
        $RequestorName,
        # Requester's email
        [Parameter(Mandatory=$true)]
        [String]
        $RequestorEmail,
        # Requester's phone number
        [Parameter(Mandatory=$true)]
        [String]
        $RequestorPhone
    )
    
    begin{
        $links = @()
        $headers = @{
            'Accept' = 'text/html, application/xhtml+xml, image/jxr, */*'
            'Referer' = "$($CA)/ca/ee/ca/profileSelect?profileID=cspMultiSANCert"
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
                'profileID' = 'cspMultiSANCert'
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
        if ((Test-Path -Path "$CSRFolderPath\links.csv" -PathType Leaf) -eq $true){
            $links | Export-Csv -Path "$CSRFolderPath\links.csv -Append"
        }
        else {
            $links | Export-Csv -Path "$CSRFolderPath\links.csv"
        }
    }
}

#############################################################################################################################################################

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

##################################################################################################################################

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

###########################################################################################################################################