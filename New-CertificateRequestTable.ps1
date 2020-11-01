function New-CertificateRequestTable {
    [CmdletBinding()]
    param (
        [Alias("CSRPath")]
        [Parameter(Mandatory=$true)]
        [String]
        $CSRFilePath,
        [Alias("SiteUrl")]
        [Parameter(Mandatory=$true)]
        [String]
        $SiteUrl,
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
        [Parameter(Mandatory=$true)]
        [Switch]
        $ResolveDNSNames
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
                    else {
                        $ip.RemoteAddress.IPAddressToString = ''
                    }
                foreach ($el in $link_csv){
                    if($el.Hostname -eq $dns_name) {
                        $linkRequest = $el.RequestID
                        $linkID = $linkRequest.TrimStart("https://$($SiteUrl)/ca/ee/ca/checkRequest?requestID=")
                    }
                }

                $obj = @(
                
                    [PSCustomObject]@{
                        CA = $CA
                        RequestID = $linkID
                        RequestLink = $linkRequest
                        Hostname = $dns_name
                        IPaddress = $ip.RemoteAddress.IPAddressToString
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
