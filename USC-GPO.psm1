function Get-GPOLinksByGuid {
[CmdLetBinding()]
Param($GUID,[switch]$DNOnly)
    # Empty array to hold all possible GPO links            
    $gPLinks = @()            
            
    # GPOs linked to the root of the domain            
    # !!! Get-ADDomain does not return the gPLink attribute            
    $gPLinks += Get-ADObject -Identity (Get-ADDomain).distinguishedName `
    -Properties name, distinguishedName, gPLink, gPOptions            
            
    # GPOs linked to OUs            
    # !!! Get-GPO does not return the gPLink attribute            
    $gPLinks += Get-ADOrganizationalUnit -Filter * -Properties name, `
    distinguishedName, gPLink, gPOptions            
            
    # GPOs linked to sites            
    $gPLinks += Get-ADObject -LDAPFilter '(objectClass=site)' `
    -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" `
    -SearchScope OneLevel -Properties name, distinguishedName, gPLink, gPOptions    
    If ($GUID) {
        $GPOLdapPath = '\[LDAP://cn={' + $GUID + '},cn=policies,cn=system,DC=usc,DC=internal.*'
        $gPLinks = $gPLinks | Where-Object { $_.gPLink -match $GPOLdapPath }
    }
    If ($DNOnly) {
        $gPLinks | ForEach-Object {
            $Links = @($_.gPLink -split {$_ -eq '[' -or $_ -eq ']'} | Where-Object {$_})
            For ($i = $Links.count -1; $i -ge 0; $i --) {
                If ($links[$i] -match $GUID) {
                    [PSCustomObject]@{
                        'DN' = $_.DistinguishedName
                        'Precedence' = $links.count - $i
                    }
                }
            }
        }
    } else {
        $gPLinks
    }
}

function Get-GPOLinksByList {
    [CmdLetBinding()]
    Param([Parameter(Mandatory=$True)]$LinkList,$GUID,[switch]$DNOnly)
    if ($GUID) {
        $GPOLdapPath = '\[LDAP://cn={' + $GUID + '},cn=policies,cn=system,DC=usc,DC=internal.*'
        $gPLinks = $LinkList | Where-Object {$_.gPLink -match $GPOLdapPath}
    } else {
        $gPLinks = $LinkList
    }
    If ($DNOnly) {
        $gPLinks | ForEach-Object {
            $Links = @($_.gPLink -split {$_ -eq '[' -or $_ -eq ']'} | Where-Object {$_})
            For ($i = $Links.count -1; $i -ge 0; $i --) {
                If ($links[$i] -match $GUID) {
                    [PSCustomObject]@{
                        'DN' = $_.DistinguishedName
                        'Precedence' = $links.count - $i
                    }
                }
            }
        }
    } else {
        $gPLinks
    }
}

function Get-GPOLinksByName {
    [CmdLetBinding()]
    Param($GPOName,$LinkList,[switch]$DNOnly)
    if ($LinkList) {
        Get-GPOLinksByList -GUID (Get-GPO -Name $GPOName).Id -LinkList $LinkList -DNOnly:$DNOnly 
    } else {
        Get-GPOLinksByGuid -GUID (Get-GPO -Name $GPOName).Id -DNOnly:$DNOnly
    }
}

function Get-SubOUList {
[CmdLetBinding()]
Param($Root)
    Begin {}
        Process {
        Write-Verbose "Searching root $Root for SubOUs"
        Get-ADOrganizationalUnit -SearchBase $Root -Filter * | Select-Object -ExpandProperty DistinguishedName
    }
}

function Find-DevGPO {
[CmdLetBinding()]
Param($ProdGPO)
    Get-GPO -Name (Get-DevGPOName -ProdGPO $ProdGPO) -ErrorAction SilentlyContinue
}

function Get-DevGPOName {
    [CmdLetBinding()]
    Param($ProdGPO)
    $ProdGPO.DisplayName -ireplace 'PROD_','DEV_'
}

function Get-DevGPONameByName {
    [CmdLetBinding()]
    Param($ProdGPOName)
    $ProdGPOName -ireplace 'PROD_','DEV_'
}

function Get-GPOLinkList {
[CmdLetBinding()]
Param($GPO) 
    Get-GPOLinksByGuid -GUID $GPO.Id
}

function Get-GPOUnsafeLinks {
[CmdLetBinding()]
Param($GPO,$ConstrainedOU='OU=MOEDev,DC=USC,DC=Internal')
    #Confirms GPO linked at a dev level
    Write-Verbose "Searching $($GPO.DisplayName) for unsafe links"
    $Result = @()
    Get-GPOLinkList -GPO $GPO | ForEach-Object {
       If ($_.DistinguishedName -notmatch $ConstrainedOU) {
           $Result = $Result + $_.DistinguishedName
       } 
    } 
    return $Result
}

function Get-ProdGPOs {
    [CmdLetBinding()]
    Param($RootOU)
    Write-Verbose "Getting GP Inheritence for Root OU $RootOU"
    Get-GPInheritance -Target $RootOU | ForEach-Object {
        $_.GpoLinks | Where-Object { $_.DisplayName -match '^PROD_.*'} 
    }
}

function Test-DestinationOU {
    [CmdLetBinding()]
    Param($DistinguishedName)
    Write-Verbose "Testing for existance of $DistinguishedName"
    Try {
        Get-ADOrganizationalUnit -SearchBase $DistinguishedName -Filter * | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Get-ProdGPOList {
    [CmdLetBinding()]
    Param([Parameter(Mandatory=$true)]$SourceRoot)
    
    $ProdOUsToCheck = ('OU=Workstations,' + $SourceRoot),('OU=Students,' + $SourceRoot),('OU=Staff,' + $SourceRoot)
    $AllOUs = $ProdOUsToCheck | ForEach-Object {Get-SubOUList -Root $_}
    Write-Verbose "Gathering status of all Production and Dev OU's"
    $AllProdGPOs = $AllOUs | ForEach-Object {Get-ProdGPOs -RootOU $_} 
    $AllProdGPOs.DisplayName | Select-Object -Unique
}

function New-DevGPODestination {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]$InputObject,
        $SourceRoot='DC=usc,DC=internal',$DestRoot='OU=moedev,DC=usc,DC=internal')
    Begin{}
    Process {
        ForEach ($Object in $InputObject) {
            [PSCustomObject]@{
                'DN' = $Object.DN -ireplace $SourceRoot,$DestRoot
                'Precedence' = $Object.Precedence
            }
        }
    }
}

function Get-DevGPOStatus {
    [CmdLetBinding()]
    Param($SourceRoot='DC=usc,DC=internal',$DestRoot='OU=moedev,DC=usc,DC=internal')
    #Refresh Link List
    $LinkList = Get-GPOLinksByGuid
    $ProdGPOs = Get-ProdGPOList -SourceRoot $SourceRoot
    $ProdGPOs | ForEach-Object {
        $ProdGPO = Get-GPO -Name $_
        $DevGPOName = Get-DevGPOName -ProdGPO $ProdGPO
        $ExistingDev = $false
        $UnsafeLinks = $Null
        Find-DevGPO -ProdGPO $ProdGPO | ForEach-Object {
            $ExistingDev = $true
            $UnsafeLinks = (Get-GPOUnsafeLinks -GPO $_ -ConstrainedOU $DestRoot)
            $LinkedOUs = Get-GPOLinksByName -GPOName $DevGPOName -LinkList $LinkList
        }
        $Destinations = Get-GPOLinksByName -GPOName $_ -LinkList $LinkList -DNOnly | New-DevGPODestination
        [PSCustomObject]@{
            'Source' = $_
            'Name' = $DevGPOName
            'DevGPOExists' = $ExistingDev
            'UnsafeLinks' = $UnsafeLinks
            'LinkedOUs' = $LinkedOUs
            'Destinations' = $Destinations
        }
    }
}

function Split-ADPath {
    [CmdLetBinding()]
    Param([Parameter(Mandatory=$True)]$Path,
    [Parameter(ParameterSetName='Parent')][switch]$Parent,
    [Parameter(ParameterSetName='Leaf')][switch]$Leaf)
    If ($Leaf) {
        #Write-Verbose "Splitting Child from $Path"
        ([regex]'^\w\w=(?<Child>\w+(\s\w+)*),.*').Match($Path).Groups | Select-Object -Last 1 -ExpandProperty Value
    } else {
        #Write-Verbose "Splitting Path from Child"
        ([regex]'^\w\w=\w+(\s\w+)*,(?<parent>\w\w=.*)').Match($Path).Groups | Select-Object -Last 1 -ExpandProperty Value
    }
}

function New-OUPath {
    [CmdLetBinding()]
    Param($OU,[switch]$Whatif)
    $Child = Split-ADPath -Path $OU -Leaf
    $Parent = Split-ADPath -Path $OU -Parent
    If (Test-DestinationOU -DistinguishedName $Parent) {
        Write-Verbose "Creating OU $Child under $Parent"
        New-ADOrganizationalUnit -Path $Parent -Name $Child -Whatif:$Whatif
    } else {
        New-OUPath $Parent
    }
}

function Reset-DevGPO {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $True,ValueFromPipelineByPropertyName=$True)]$Source,
        [Parameter(Mandatory = $True,ValueFromPipelineByPropertyName=$True)]$Name,
        [switch]$Whatif
    )

    Begin{}
    Process {
        $SourceGPO = Get-GPO -Name $Source
        Find-DevGPO -ProdGPO $SourceGPO | ForEach-Object {
            Write-Verbose "Removing GPO $($_.DisplayName)"
            $_ | Remove-GPO
        }
        Write-Verbose "Duplicating GPO $($SourceGPO.DisplayName) to new GPO $(Get-DevGPOName -ProdGPO $SourceGPO)"
        $SourceGPO | Copy-GPO -TargetName (Get-DevGPOName -ProdGPO $SourceGPO) -CopyAcl -WhatIf:$Whatif
    }
}

function Get-UnexcludedDestinationOU {
    [CmdLetBinding()]
    Param($DestinationOU,$Exclusions)

    $DestIsValid = $true
    $Exclusions | ForEach-Object {
        If ($DestinationOU.DN -match $_) {
            $DestIsValid = $False
        }
    }
    If ($DestIsValid){
        $DestinationOU
    }

}

<#function New-DevGPOLink {
    [CmdLetBinding()]
    Param([Parameter(Mandatory=$true)]$DevGPO,$DestinationOU,
    [switch]$Whatif)
    Begin{}
    Process {
        If (-Not (Test-DestinationOU -DistinguishedName $DestinationOU)) {
            Write-Verbose "Creating new OU $DestinationOU"
            New-OUPath -OU $DestinationOU -Whatif:$Whatif 
        }
        Write-Verbose "Creating new GPLink of $($DevGPO.DisplayName) to $DestinationOU"
        $DevGPO | New-GPLink -Target $DestinationOU -LinkEnabled Yes -WhatIf:$Whatif
    }
} #>

function Reset-DevGPOs {
    [CmdLetBinding()]
    Param([switch]$Whatif,[Parameter(Mandatory=$true,ValueFromPipeline=$True)][array]$InputObject,[int]$Wait=5)

     $InputObject | ForEach-Object {
         $_ | Reset-DevGPO -Whatif:$Whatif
     }
     Start-Sleep -Seconds $Wait
}

function New-DevGPOLink {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]$InputObject,
        [Parameter(Mandatory=$True)]$Destinations,
        $Exclusions = @('Workstations - Custom','Utility','Servers'),
        [switch]$Whatif
    )

    Begin{
        $DestinationOUs = $Destinations | ForEach-Object {
            Get-UnexcludedDestinationOU -DestinationOU $_ -Exclusions $Exclusions 
        }
    }

    Process {
        ForEach ($DestinationOU in $DestinationOUs){
            If (-Not (Test-DestinationOU -DistinguishedName $DestinationOU.DN)) {
                Write-Verbose "Creating new OU $($DestinationOU.DN)"
                New-OUPath -OU $DestinationOU.DN -Whatif:$Whatif 
            }
            Write-Verbose "Creating new GPLink of $($DevGPO.DisplayName) to $($DestinationOU.DN)"
            $InputObject | New-GPLink -Target $DestinationOU.DN -LinkEnabled Yes -WhatIf:$Whatif
        }
    }
}

#Get-DevGPOStatus | ForEach-Object { $_ | Reset-DevGPOs -Wait 0 -Verbose | New-DevGPOLink -Destinations $_.Destinations -Verbose }
<#Get-DevGPOStatus | ForEach-Object { 
    Get-GPO -Name (Get-DevGPONameByName $_.Source) | New-DevGPOLink -Destinations $_.Destinations -Verbose 
}
$AllPolicies = Get-DevGPOStatus
#>