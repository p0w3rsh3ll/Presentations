
Function Get-MSIZapInfo {
<#
  
.SYNOPSIS    
    Get MSIinstaller related info from the registry
  
.DESCRIPTION  
    Get MSIinstaller related info from the registry
  
.PARAMETER ShowSupersededPatches
    Switch to show supersed patches of MSI products instead of default installed MSI products
  
.NOTES    
    Name: Get-MSIZapInfo
    Author: Emin Atac
    DateCreated: 02/01/2012
  
.LINK    
    https://p0w3rsh3ll.wordpress.com
  
.EXAMPLE    
    Get-MSIZapInfo | fl -property *
    Pipe it into format-list and show all the returned properties
  
.EXAMPLE    
    Get-MSIZapInfo | fl -property Displayname
    Pipe it into format-list and show only the displayname of MSI installed products
  
.EXAMPLE
    Get-MSIZapInfo | Format-Custom -Depth 2
    Pipe it into format-custom to explore its depth
  
.EXAMPLE    
    Get-MSIZapInfo | Select-Object -Property Displayname,Displayversion,Publisher | ft -AutoSize -HideTableHeaders    
    Select some properties and pipe it into the format-table cmdlet
  
.EXAMPLE
    Get-MSIZapInfo -ShowSupersededPatches |  ft -AutoSize -Property LocalPackage,Superseded,DisplayName
    Show superseded patches of all MSI products installed and filter some of its properties using the format-table cmdlet
  
.EXAMPLE
    (Get-MSIZapInfo | Where-Object -FilterScript  {$_.Displayname -match "LifeCam"}) | 
    fl -Property *
    Find a specific MSI product installed by its displayname and show all its properties
  
.EXAMPLE
    (Get-MSIZapInfo | Where-Object -FilterScript  {$_.Displayname -match "Silverlight"}).AllPatchesEverinstalled |
     ft -AutoSize
    Look for Silverlight and show all its patches ever installed
  
.EXAMPLE
    Get-MSIZapInfo | Where-Object -FilterScript  {$_.Displayname -match "Microsoft Office" } | 
    fl -Property Displayname,RegistryGUID,UninstallString,ConvertedGUID
  
.EXAMPLE    
    $sb=[scriptblock]::Create((Get-Command Get-MSIZapInfo).Definition)
    Invoke-Command -ComputerName RemoteComputername -ScriptBlock $sb
    List all MSI products of a remote computer
  
.EXAMPLE        
    $sb=[scriptblock]::Create((Get-Command Get-MSIZapInfo).Definition)
    (Invoke-Command -ComputerName RemoteComupterName -ScriptBlock $sb) | % {$_.AllPatchesEverinstalled}
    List all MSI products patches ever installed on a remote computer
  
.EXAMPLE
    Get-MSIZapInfo | % {
        if ($null -ne $_.AllPatchesEverinstalled) {
            $_.AllPatchesEverinstalled | 
            Where-Object {$_.Superseded -eq $false} |  
            ft -AutoSize -Property LocalPackage,Superseded,DisplayName 
        }
    }    
    Show non supersed patches of all installed MSI products
  
#>
[CmdletBinding()]
Param(
    [Parameter()]
    [switch]$ShowSupersededPatches
)
Begin {  
  
    # idea from https://adameyob.com/2016/03/convert-program-guid-product-code/
    Function Convert-RegistryGUID {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateLength(32,32)]
        [string]$String
    )
    Begin {}
    Process {
        $raw = [regex]::replace($String, '[^a-zA-Z0-9]','')
        (
            [guid](
                -join (
                    7,6,5,4,3,2,1,0,11,10,9,8,15,14,13,12,17,16,19,18,21,20,23,22,25,24,27,26,29,28,31,30,32 |
                    ForEach-Object {
                        $raw[$_]
                    }
                )
            )
        ).Tostring('D').ToUpper()
    }
    End{}
    }
    # Convert-RegistryGUID "00004109110000000000000000F01FEC"
}
Process {  
    # Main: read info in the registry and populate array with object that have the properties we are looking for
  
    if (Test-Path HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData)
    {
        $root = Get-Childitem HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData
  
        # Initialize the patches and products global array
        $patchesar = @()
        $prodcutssar = @()
  
        # Cycle through all SIDs
        foreach ($i in $root)
        {
            Write-Verbose -Message "Dealing with $($i.Name)" #-Verbose:$true
  
            # Convert the SID to an account name
            $UserWhoInstalled = ConvertTo-NtAccount $i.PSChildName
  
            # Build the main subkey
            $subkey = Join-Path -Path $i.PSParentPath -ChildPath $i.PSChildName
  
  
            # Get the whole list of patches
            if (Test-Path "$subkey\Patches")
            {
                $patches =  Get-Childitem "$subkey\Patches"
                foreach ($k in $patches)
                {
                    $patchkey = Join-Path -Path $k.PSParentPath -ChildPath $k.PSChildName
  
                    $PatchObject = New-Object -TypeName PSObject -Property @{
                        RegistryGUID = $k.PSChildName
                        InstalledBy = $UserWhoInstalled
                        LocalPackage = (Get-ItemProperty -Path $patchkey).LocalPackage
                        }
  
                    # Add our object to the global array
                    $patchesar += $PatchObject
                }
            }
  
            # Get the list of products and their properties
            if (Test-Path "$subkey\Products")
            {
                $products = Get-Childitem "$subkey\Products"
  
                foreach ($j in $products)
                {
                    Write-Verbose -Message "Dealing with $($j.Name)" # -Verbose:$true
  
                    # Build the subkey and gather all properties
                    $productkey = Join-Path -Path $j.PSParentPath -ChildPath $j.PSChildName
                    $productInstallProperties = Get-ItemProperty -Path $productkey\InstallProperties
  
                    # Populate our object with all the properties we are interested in
                    $ProductObj = New-Object -TypeName PSObject -Property @{  
                        InstalledBy = $UserWhoInstalled
                        RegistryGUID = $j.PSChildName
                        Displayname = $productInstallProperties.DisplayName
                        Publisher = $productInstallProperties.Publisher
                        DisplayVersion = $productInstallProperties.DisplayVersion
                        InstallDate = ([System.DateTime]::ParseExact($productInstallProperties.InstallDate,"yyyyMMdd",[System.Globalization.CultureInfo]::InvariantCulture))
                        LocalPackage = $productInstallProperties.LocalPackage
                        UninstallString = ($productInstallProperties.UninstallString -replace "msiexec\.exe\s/[IX]{1}","")
                        ConvertedGUID = (Convert-RegistryGUID -String $j.PSChildName)
                        }
  
                    # Get the list of patches GUID for a product
  
                    # Build the array of current patches w/o superseded patches from the Allpatches value found in the registry
                    $AllpatchesValuear = @()
                    $AllpatchesValue = (Get-ItemProperty -Path $productkey\Patches).Allpatches -split "`n"
                    for ($i = 0 ; $i -le ($AllpatchesValue.Count - 1) ; $i++)
                    {
                        $AllpatchesValuear += $AllpatchesValue[$i]
                    }
  
                    # Cycle to all patches found by reading subkeys
                    $Allproductpatches = Get-Childitem "$productkey\Patches"
                    $AllpatchesInstalled = @()
                    if ($Allproductpatches -ne $null)
                    {
                        foreach ($l in $Allproductpatches)
                        {
                            $patchesubkey = Join-Path -Path $l.PSParentPath -ChildPath $l.PSChildName
                            $patchproperties = Get-ItemProperty -Path $patchesubkey
  
                            $PatchObj = New-Object -TypeName PSObject -Property @{            
                                RegistryGUID = $l.PSChildName
                                Displayname =$patchproperties.DisplayName
                                InstallDate = ([System.DateTime]::ParseExact($patchproperties.Installed,"yyyyMMdd",[System.Globalization.CultureInfo]::InvariantCulture))
                                }
  
                            # Prepare to define a supersedence property
                            switch ($patchproperties.State)
                            {
                                2 { $Superseded = $true}
                                1 { $Superseded = $false}
                                default { $Superseded = "Unknown"}
                            }                    
                            $PatchObj | add-member Noteproperty -Name Superseded -Value $Superseded
  
                            if (Test-MatchFromHashTable -array $patchesar -testkey RegistryGUID -string $PatchObj.RegistryGUID)
                            {
                               $PatchObj | add-member Noteproperty -Name LocalPackage -Value (Get-MatchFromHashTable -array $patchesar -testkey RegistryGUID -string $PatchObj.RegistryGUID).LocalPackage
                            }
                            # Add to array
                            $AllpatchesInstalled += $PatchObj
                        }
                    }
  
                    $ProductObj | add-member Noteproperty -Name AllPatchesEverinstalled -Value $AllpatchesInstalled
                    $ProductObj | add-member Noteproperty -Name AllCurrentPatchesinstalled -Value $AllpatchesValuear
  
                    # Add our object to the global array
                    $prodcutssar += $ProductObj
  
                } # end of foreach
            } # end of if test-path products
        } # end of foreach root
    } # end of test-path
  
    if ($ShowSupersededPatches)
    {
        # Superseded patches
        $Supersededpatches = @()
        foreach ($j in $prodcutssar)
        {
            if ($j.AllPatchesEverinstalled -ne $null)
            {
                $Supersededpatches += ($j.AllPatchesEverinstalled | Where-Object {$_.Superseded -eq $true})
  
            }
        }
        return $Supersededpatches
    } else {
        return $prodcutssar
    }
}
End {}
}