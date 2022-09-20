# Module for Fresh Service Catalog functions
# Les Newbigging 2022
#
# See https://api.freshservice.com/#service-catalog

function Get-FreshServiceCatalogItem {
    <#
    .SYNOPSIS
        Retrieves a list of Fresh Service Catalog Items
    .DESCRIPTION
        Retrieves a list of Fresh Service Catalog Items, or a more detailed single item
    .EXAMPLE
        Get-FreshServiceCatalogItem
        Lists all available Service Catalog items.
    .EXAMPLE
        Get-FreshServiceCatalogItem -DisplayID 5
        Retrieves all the details for the catalog item with display_id 5, including custom_fields and child_items
    .EXAMPLE
        Get-FreshServiceCatalogItem -CategoryID 10
        Retrieves a list of items classed under category 10.
    .EXAMPLE
        Get-FreshServiceCatalogItem | Where-Object name -like "Microsoft*" | Get-FreshServiceCatalogItem
        Will give more details about all items whose names begin with 'Microsoft'
    .NOTES
        Retrieving a list will not include custom_fields and child_items, but querying for an individual catalog item will provide these details, but by piping (i.e. Get-FreshServiceCatalogItem | Get-FreshServiceCatalogItem), this will retrieve all details for every item.
        This, however, will also result in repetitive calls to the Fresh API - one per page of a hundred items, plus one per item, which may result in a slowing of response due to throttling (see https://api.freshservice.com/#rate_limit).
    .INPUTS
        Object. Must include display_id as a property.
    .OUTPUTS
        Object[]. Object representations of each catalog item.
    #>
    [CmdletBinding(DefaultParameterSetName='service_items')]
    param(
        # Display ID of the item
        [parameter(Mandatory=$true,
            ParameterSetName='service_item',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('display_id')]
        [int64]$DisplayID,

        # Category ID of items to return
        [parameter(Mandatory=$false,
            ParameterSetName='service_items')]
        [Alias('category_id')]
        [int64]$CategoryID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions

    $path = "service_catalog/items"
    $field = $PSCmdlet.ParameterSetName
    if ($DisplayID)
    {
        $path += "/$DisplayID"
    } else {
        $path += "?per_page=100"

        if ($CategoryID)
        {
            $path += "&category_id=$CategoryID"
        }
    }
    try {
        Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Search-FreshServiceCatalogItem {
    <#
    .SYNOPSIS
        Searches for items that match the search term
    .DESCRIPTION
        Searches for service catalog items that match the search term
    .NOTES
        Unlike Get-FreshServiceCatalogItem, the results include the child_items and custom_fields properties by default.
    .INPUTS
        None. This does not accept pipeline input.
    .OUTPUTS
        Object[]. An array ofobjects with the service catalog items that meet the search criteria.
    .EXAMPLE
        Search-FreshServiceCatalogItem -SearchTerm "Laptop"
        Returns all service catalog items that include 'Laptop' in their name or description.
    #>
    [CmdletBinding(DefaultParameterSetName='service_items')]
    param(
        # The search term to find the service catalog items
        [parameter(Mandatory=$true)]
        [Alias('search_term')]
        [string]$SearchTerm,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions

    try {
        Invoke-FreshAPIGet -path "service_catalog/items/search?per_page=100&search_term=$SearchTerm" -field "service_items" -system $System -verbose:$verbosity | Select-Object *,@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Get-FreshServiceCategory {
    <#
    .SYNOPSIS
        Retrieves a list of Fresh Service Catalog categories
    .DESCRIPTION
        Lists all service categories in your Freshservice service desk.
    .EXAMPLE
        Get-FreshServiceCategory | Select-Object id,name,description
        id name               description
        -- ----               -----------
        10 Hardware           A list of available hardware
        11 Software           A list of available software
        12 Cloud applications A list of available cloud applications
    .INPUTS
        None. Does not accept pipeline input.
    .OUTPUTS
        Object[]. An array of object representations of the catalog categories.
    #>
    [CmdletBinding()]
    param(
        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions

    try {
        Invoke-FreshAPIGet -path "service_catalog/categories" -field "service_categories" -system $System -verbose:$verbosity | Select-Object *,@{name='category_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

