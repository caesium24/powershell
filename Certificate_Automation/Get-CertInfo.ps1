function Get-DomainComp {Param($OU) get-ADComputer -Filter {ObjectClass -eq "Computer"} | Where-Object {$_.DistinguishedName -like "*$OU*"}}

function Get-CertInfo {
    Param(
        [Parameter(Mandatory=$true)]
        [String]
        $SavePath,
        [Parameter(Mandatory=$true)]
        [String]
        $OU,
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]
        $Credentials
    )
    
    $comps = Get-DomainComp($OU)
    $thumbprints = @()

    foreach ($comp in $comps) {
        $result = Invoke-Command -ComputerName $comp.Name -Credential $creds -ScriptBlock {
            $certificate = Get-ChildItem "Cert:\LocalMachine\My"
            $obj = @(
                [PSCustomObject]@{
                    Hostname = $env:COMPUTERNAME
                    Thumbprint = $certificate.Thumbprint
                    Expiration = $certificate.NotAfter
                }
            )
            $obj
        }
        $thumbprints += $result
    }
    $thumbprints | Export-Csv -Path $SavePath
}
