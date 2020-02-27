# The Get-VDITemplate function allows you to grab all active Citrix Machine catalogs and store them in the $templates variable.
function Get-VDITemplate {
    Add-PSSnapin Citrix*
    $templates = @()
    Get-ProvScheme |
    ForEach-Object { $template = @(
        [PSCustomObject]@{
        TemplateName = $_.ProvisioningScheme
        TemplateUID = $_.ProvisioningSchemeID
        TemplateImagePath = $_.MasterImageVM
        Count = $_.MachineCount
        Type = if($_.CleanOnBoot -eq $true) {"Ephemereal"} else {"Static"}
        }
    )
    $templates += $template
    }
    return $templates
}

# The Set-VDIImage function pulls the latest snapshots of the machine catalogs given in the argument and updates their images.
function Set-VDIImage ($templates) {
    foreach ($template in $templates) {
        if ($template.Type -eq "Ephemereal") {
            $path = $template.TemplateImagePath
            $num = $path.IndexOf(".vm") + 4
                $path = $path.Remove($num)
            $newSnap = Get-ChildItem -Recurse -Path $path
            Publish-ProvMasterVmImage -ProvisioningSchemeName $template.TemplateName -MasyerImageVM $newSnap[-1].FullPath -RunAsynchronously
        }
    }
}

# The Update-VDIImage function calls the above to functions and allows you to select all Machine Catalog Images or create a
# custom list to be updated.

function Update-VDIImage {
    [CmdletBinding()]
    param (
    [Parameter(Mandatory=$true)]
    [Switch[]]
    $All
    )
    $templates = Get-VDITemplate
    if ($All -eq $true) {
        Set-VDIImage($templates)
    }
    else {
        $TmplList = @()
        $templates.TemplateName
        Write-Host -ForegroundColor Green "From the above list of VM Template names, please type in which you would like to update an image.
        Example: web.client.com, test.client.com"
        $List = Read-Host "Please enter Template Names"
        $List = $List.split(',')
        foreach ($name in $List) {
            $name.TrimStart(' ')
            $TmplList += $name
        }
        foreach ($template in $templates) {
            if ($TmplList -contains $template.TemplateName){
                Set-VDIImage($template)
            }
        }
    }
}
