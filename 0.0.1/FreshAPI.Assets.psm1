# Module for Fresh Assets functions
#
# See https://api.freshservice.com/#assets
function Get-FreshAsset {
    <#
    .SYNOPSIS
        Retrieves asset(s) from Fresh
    .DESCRIPTION
        Retrieves assets from Fresh, by ID, by filter conditions, or all.
    .EXAMPLE
        Get-FreshAsset
        Returns all fresh assets
    .EXAMPLE
        Get-FreshAsset -AssetId 42
        Returns the details for asset with a display ID of 42
    .EXAMPLE
        Get-FreshAsset -UserId 4 -AssetState InUse  
        Returns all assets assigned to user 4 that are in use
    .EXAMPLE
        Get-FreshAsset -Tag 'critical' -IncludeTypeFields 
        Retrieves all assets with the 'critical' tag, and returns the 'type_fields' along with the assets
    .OUTPUTS
        Object[]
    .NOTES
        Includes textualised properties for impact, priority, source, status & urgency for readability.
    #>
    [cmdletBinding(DefaultParameterSetName='assets')]
    param(
        # Finds a single Fresh Asset by (numeric) ID
        [parameter(Mandatory=$true,
            ParameterSetName='asset',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('asset_id')]
        [int64]$AssetId,

        # Search by ID of the asset type of the asset
        [parameter(ParameterSetName='assets')]
        [Alias('asset_type_id')]
        [int64]$AssetTypeId,

        # Search by ID of the department to which the asset has been assigned
        [parameter(ParameterSetName='assets')]
        [Alias('department_id')]
        [int64]$DepartmentId,

        # Search by ID of the location to which the asset has been assigned
        [parameter(ParameterSetName='assets')]
        [Alias('location_id')]
        [int64]$LocationId,

        # Search by ID of the user to which the asset has been assigned
        [parameter(ParameterSetName='assets')]
        [Alias('user_id')]
        [int64]$UserId,

        # Search by ID of the agent who manages the asset
        [parameter(ParameterSetName='assets')]
        [Alias('agent_id')]
        [int64]$AgentId,

        # Search by the name of the asset
        [parameter(ParameterSetName='assets')]
        [Alias('name')]
        [int64]$AssetName,
        
        # Search by state of the asset
        [parameter(ParameterSetName='assets')]
        [Alias('asset_state')]
        [string]$AssetState,

        # Search for tickets with a particular tag 
        [parameter(ParameterSetName='assets')]
        [string]$Tag,

        # Include Type Fields with the tickets
        [switch]$IncludeTypeFields,

        # Which Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment
    )
    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this enmsure the -Verbose setting is passed through to subsequent functions
    }

    process {
        $path = "assets"
        $field = $PSCmdlet.ParameterSetName
        write-verbose "ParameterSetName: $field"

        $FilterParameters = @()

        switch($field)
        {
            'asset' {
                $path += "/$AssetId"
            }
        }

        $Includes = @()
        if ($IncludeTypeFields)
        {
            $Includes += "type_fields"
        }

        if ($Includes.count -ne 0)
        {
            $FilterParameters += "include=" + ($Includes -join ',')
        }

        $Queries = @()

        if ($AssetTypeId)
        {
            $Queries += "asset_type_id:$AssetTypeId"
        }

        if ($DepartmentId)
        {
            $Queries += "department_id:$DepartmentId"
        }

        if ($LocationId)
        {
            $Queries += "location_id:$LocationId"
        }

        if ($UserId)
        {
            $Queries += "user_id:$UserId"
        }

        if ($AssetName)
        {
            $Queries += "name:'$AssetName'"
        }

        if ($AgentID)
        {
            $Queries += "agent_id:$AgentID"
        }

        if ($AssetState)
        {
            $StateArray = @()
            foreach ($s in $AssetState)
            {
                $StateArray += "asset_state:$s)"
            }
            $Queries += "(" + ($StateArray -join ' OR ') + ")"
        }    

        if ($Tag)
        {
            $Queries += "tag:'$Tag'"
        }

        if ($Queries.count -ne 0)
        {
            $FilterParameters += 'filter="' + ($Queries -join ' AND ') + '"'
        }

        if ($FilterParameters.count -ne 0)
        {
            $path += "?" + ($FilterParameters -join '&')
        }
        write-verbose "Path: $path"

        try {
            Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='asset_id';exp={$_.display_id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Get-FreshAssetType {
    <#
    .SYNOPSIS
        Retrieves asset type(s) from Fresh
    .DESCRIPTION
        Retrieves asset types from Fresh, by ID or all.
    .EXAMPLE
        Get-FreshAssetType
        Returns all fresh asset types
    .EXAMPLE
        Get-FreshAssetType -AssetTypeId 42
        Returns the details for asset type with an ID of 42
    .EXAMPLE
        Get-FreshAssetType
        Returns all asset types in the system
    .OUTPUTS
        Object[]
    #>
    [cmdletBinding(DefaultParameterSetName='asset_types')]
    param(
        # Finds a single Fresh Asset Type by (numeric) ID
        [parameter(Mandatory=$true,
            ParameterSetName='asset_type',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('asset_type_id')]
        [int64]$AssetTypeId,

        # Which Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment
    )
    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this enmsure the -Verbose setting is passed through to subsequent functions
    }

    process {
        $path = "asset_types"
        $field = $PSCmdlet.ParameterSetName
        write-verbose "ParameterSetName: $field"

        switch($field)
        {
            'asset_type' {
                $path += "/$AssetTypeId"
            }
        }

        write-verbose "Path: $path"

        try {
            Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='asset_type_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}
