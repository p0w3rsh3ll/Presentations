
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
  
    if (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData' -PathType Container) {
        
  
        # Initialize the patches and products global array
        $patchesar = @()
        $prodcutssar = @()
  
        # Cycle through all SIDs
        Get-Childitem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData' | 
        ForEach-Object {
        
            $i = $_
            
            Write-Verbose -Message "Dealing with $($i.Name)"
  
            # Convert the SID to an account name
            $UserWhoInstalled = $(
                try {
                    (New-Object System.Security.Principal.SecurityIdentifier -ArgumentList "$($i.PSChildName)").translate(
                    [System.Security.Principal.NTAccount]
                    ).Value
                } catch {
                    "$($i.PSChildName)"
                }
            )
  
            # Build the main subkey
            $subkey = Join-Path -Path $i.PSParentPath -ChildPath $i.PSChildName
  
            # Get the whole list of patches
            if (Test-Path -Path "$subkey\Patches" -PathType Container) {
                
                Get-Childitem -Path "$subkey\Patches" | 
                ForEach-Object {
                    $k = $_
                    $patchesar += [PSCustomObject]@{
                        RegistryGUID = $k.PSChildName
                        InstalledBy = $UserWhoInstalled
                        LocalPackage = (Get-ItemProperty -Path (Join-Path -Path $k.PSParentPath -ChildPath $k.PSChildName)).LocalPackage
                    }
                }
            }
  
            # Get the list of products and their properties
            if (Test-Path -Path "$subkey\Products" -PathType Container) {
            
                Get-Childitem  -Path "$subkey\Products" | 
                ForEach-Object {
                    $j = $_

                    Write-Verbose -Message "Dealing with $($j.Name)" # -Verbose:$true
  
                    # Build the subkey and gather all properties
                    $productkey = Join-Path -Path $j.PSParentPath -ChildPath $j.PSChildName

                    if (Test-Path -Path "$($productkey)\InstallProperties" -PathType Container) {
                    
                        $ipr = Get-ItemProperty -Path $productkey\InstallProperties
  
                        # Populate our object with all the properties we are interested in
                        $productssar += [PSCustomObject]@{  
                            InstalledBy = $UserWhoInstalled
                            RegistryGUID = $j.PSChildName
                            Displayname = $ipr.DisplayName
                            Publisher = $ipr.Publisher
                            DisplayVersion = $ipr.DisplayVersion
                            InstallDate = ([System.DateTime]::ParseExact($ipr.InstallDate,'yyyyMMdd',[System.Globalization.CultureInfo]::InvariantCulture))
                            LocalPackage = $ipr.LocalPackage
                            UninstallString = ($ipr.UninstallString -replace 'msiexec\.exe\s/[IX]{1}','')
                            ConvertedGUID = (Convert-RegistryGUID -String $j.PSChildName)
                            AllCurrentPatchesInstalled = $(
                                @(
                                    (Get-ItemProperty -Path $productkey\Patches).Allpatches -split "`n" | 
                                    ForEach-Object {
                                        $_
                                    }
                                )
                            ) ; # AllCurrentPatchesInstalled
                            AllPatchesEverInstalled = $(
                                # Cycle to all patches found by reading subkeys
                                Get-Childitem -Path "$productkey\Patches" | 
                                ForEach-Object {
                                    $l = $_
                    
                                    $pap = Get-ItemProperty -Path (Join-Path -Path $l.PSParentPath -ChildPath $l.PSChildName)
  
                                    [PSCustomObject]@{            
                                        RegistryGUID = $l.PSChildName
                                        Displayname =$pap.DisplayName
                                        InstallDate = ([System.DateTime]::ParseExact($pap.Installed,'yyyyMMdd',[System.Globalization.CultureInfo]::InvariantCulture))
                                        Superseded =$( 
                                            Switch ($pap.State) {
                                                2 { $true  ; break }
                                                1 { $false ; break }
                                                default {'Unknown'}
                                            }
                                        )
                                        LocalPackage = $(
                                            ($patchesar | Where-Object { $_.RegistryGUID -eq $l.PSChildName }).LocalPackage
                                        )
                                    }
                                }

                            ) # AllPatchesEverInstalled
                        }
 
                    } else {
                        Write-Warning -Message "$($productkey)\InstallProperties not found"
                    }  
                } # end of foreach
            } # end of if test-path products
        } # end of foreach root
    } # end of test-path
  
    if ($ShowSupersededPatches) {
        # Superseded patches
        $productssar | 
        ForEach-Object {
            if ($null -ne $_.AllPatchesEverinstalled) {
                $_.AllPatchesEverinstalled | Where-Object {$_.Superseded}
            }
        }
    } else {
        $productssar
    }
}
End {}
}