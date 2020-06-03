########################################################################################################################################################################
<#

Creates new SSL Certificate Requests and Keys using either vCenter, CSV, or user input to populate hostnames

    Version : 1.6
    Author  : Eshton Brogan & Sid Johnson
    Created : 09 October 2019
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
        $openssl = where.exe /R 'C:\Program Files' openssl.exe | Select-Object -First 1
            if ($null -eq $openssl){
                Write-Warning -Message "Openssl not found in C:\Program Files directory. Searching again in C:\"
                $openssl = where.exe /R 'C:\' openssl.exe | Select-Object -First 1
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
