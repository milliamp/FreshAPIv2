# Module for Fresh Locations functions
# Les Newbigging 2022
# 
# See https://api.freshservice.com/#locations

function New-FreshLocation {
    <#
    .SYNOPSIS
        This operation allows you to create a new location in Freshservice.
    .DESCRIPTION
        This operation allows you to create a new location in Freshservice.
    .INPUTS
        None. This function does not accept pipeline input.
    .OUTPUTS
        Object. A representation of the newly created location.
    .EXAMPLE
        New-FreshLocation -Name "The Gherkin" -Line 1 "40th floor" -Line2 "30 St Mary Axe" -City "London" -Country "United Kingdom" -Zipcode "EC3"
        Creates a new location
    #>
    [CmdletBinding()]
    param(
        # Name of the location.
        [parameter(Mandatory=$true)]
        [string]$Name,

        # ID of the parent location
        [Alias('parent_location_id')]
        [int64]$ParentLocationID,

        # User ID of the primary contact (must be a user in Freshservice)
        [Alias('primary_contact_id')]
        [int64]$PrimaryContactID,

        # Address line 1.
        [string]$Line1,
        
        # Address line 2.
        [string]$Line2,
        
        # Name of the Country.
        [string]$City,

        # Name of the State.
        [string]$State,
        
        # Name of the Country.
        [string]$Country,

        # Zipcode of the location.
        [string]$Zipcode,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    $Address = @{
        line1 = $Line1
        line2 = $Line2
        city = $City
        state = $State
        country = $Country
        zipcode = $Zipcode            
    }

    # find empty keys in address
    $EmptyKeys = @()
    foreach ($key in $Address.keys)
    {
        if ($null -eq $Address[$key] -or $Address[$key] -eq '' -or $Address[$key] -eq 0)
        {
            $EmptyKeys += $key
        }
    }

    # remove empty keys
    foreach ($key in $EmptyKeys)
    {
        $Address.Remove($key)
    }
    
    if ($Address.Keys.count -eq 0)
    {
        $Address = $null
    }

    $Body = @{
        name = $Name
        parent_location_id = $ParentLocationID
        primary_contact_id = $PrimaryContactID
        address = $Address
    }

    # find empty keys
    $EmptyKeys = @()
    foreach ($key in $Body.keys)
    {
        if ($null -eq $Body[$key] -or $Body[$key] -eq '' -or $Body[$key] -eq 0)
        {
            $EmptyKeys += $key
        }
    }

    # remove empty keys
    foreach ($key in $EmptyKeys)
    {
        $Body.Remove($key)
    }

    $jsonBody = $Body | ConvertTo-Json

    try {
        Invoke-FreshAPIPost -path "locations" -field "location" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')  | Select-Object *,@{name='location_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Get-FreshLocation {
    <#
    .SYNOPSIS
        View information about locations in Fresh.
    .DESCRIPTION
        View information about locations in Fresh.
    .INPUTS
        Object. Requires the location_id property.
    .OUTPUTS
        Object[]. Representation(s) of the location(s) returned.
    .EXAMPLE
        Get-FreshLocation
        Returns all the locations in Fresh
    .EXAMPLE
        Get-FreshLocation -LocationID 123
        Retrieves the details of location 123 from Fresh.
    .EXAMPLE
        Get-FreshLocation -Name "The Gherkin"
        Retrieves details for the location with the name 'The Gherkin'
    #>
    [CmdletBinding(DefaultParameterSetName='locations')]
    param(
        # ID of Fresh location
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true,
            ParameterSetName='location')]
        [Alias('location_id')]
        [int64]$LocationID,

        # Name of the location.
        [parameter(Mandatory=$false,
            ParameterSetName='locations')]
        [string]$Name,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions

    $path = "locations"
    $field = $PSCmdlet.ParameterSetName

    if ($LocationID)
    {
        $path += "/$LocationID"
    } else {
        $path += "?per_page=100"

        if ($Name)
        {
            $path += '&query="name:' + "'$Name'" + '"'
        }
    }

    try {
        Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='location_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Update-FreshLocation {
    <#
    .SYNOPSIS
        This operation allows you to update an existing location.
    .DESCRIPTION
        This operation allows you to update an existing location.
    .INPUTS
        Object. Must include the location_id property.
    .OUTPUTS
        Object. A representation of the updated location.
    .EXAMPLE
        Update-FreshLocation -LocationID 123 -Country "United Kingdom"
        Modifies the country property of the Fresh location with ID 123 
    .EXAMPLE
        Get-FreshLocation -Name "The Shard" | Update-FreshLocation -City "London"
        Updates the location called "The Shard" setting 'City' in the address to 'London'
    #>
    [CmdletBinding()]
    param(
        # Unique ID of the location
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('location_id')]
        [int64]$LocationID,

        # Name of the location.
        [string]$Name,

        # ID of the parent location
        [Alias('parent_location_id')]
        [int64]$ParentLocationID,

        # User ID of the primary contact (must be a user in Freshservice)
        [Alias('primary_contact_id')]
        [int64]$PrimaryContactID,

        # Address line 1.
        [string]$Line1,
        
        # Address line 2.
        [string]$Line2,
        
        # Name of the Country.
        [string]$City,

        # Name of the State.
        [string]$State,
        
        # Name of the Country.
        [string]$Country,

        # Zipcode of the location.
        [string]$Zipcode,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )
    
    process {
        $Address = @{
            line1 = $Line1
            line2 = $Line2
            city = $City
            state = $State
            country = $Country
            zipcode = $Zipcode            
        }

        # find empty keys in address
        $EmptyKeys = @()
        foreach ($key in $Address.keys)
        {
            if ($null -eq $Address[$key] -or $Address[$key] -eq '' -or $Address[$key] -eq 0)
            {
                $EmptyKeys += $key
            }
        }

        # remove empty keys
        foreach ($key in $EmptyKeys)
        {
            $Address.Remove($key)
        }
        
        if ($Address.Keys.count -eq 0)
        {
            $Address = $null
        }

        $Body = @{
            name = $Name
            parent_location_id = $ParentLocationID
            primary_contact_id = $PrimaryContactID
            address = $Address
        }

        # find empty keys
        $EmptyKeys = @()
        foreach ($key in $Body.keys)
        {
            if ($null -eq $Body[$key] -or $Body[$key] -eq '' -or $Body[$key] -eq 0)
            {
                $EmptyKeys += $key
            }
        }

        # remove empty keys
        foreach ($key in $EmptyKeys)
        {
            $Body.Remove($key)
        }

        $jsonBody = $Body | ConvertTo-Json

        try {
            Invoke-FreshAPIPut -path "locations/$LocationID" -field "location" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')  | Select-Object *,@{name='location_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Remove-FreshLocation {
    <#
    .SYNOPSIS
        This operation allows you to delete a particular location.
    .DESCRIPTION
        This operation allows you to delete a particular location.
    .INPUTS
        Object. Must include the location_id property.
    .OUTPUTS
        None.
    .EXAMPLE
        Remove-FreshLocation -LocationID 123
        Deletes location 123 from Fresh
    #>
    [cmdletBinding()]
    param(
        # ID of Fresh location
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('location_id')]
        [int64]$LocationID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment          
    )

    process {
        try {
            Invoke-FreshAPIdelete -path "locations/$LocationID" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')  
        } catch {
            $_ | Convert-FreshError
        }
    }
}