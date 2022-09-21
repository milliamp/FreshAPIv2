# Module for Fresh Custom Object functions
# Les Newbigging 2022
#
# See https://api.freshservice.com/#custom-objects

function Get-FreshCustomObject {
    <#
    .SYNOPSIS
        Lists all the Custom objects that are present in the account, or gives you the details of a given Custom object
    .DESCRIPTION
        Lists all the Custom objects that are present in the account, or gives you the details of a given Custom object, such as field names, options for dropdown fields etc.
    .NOTES
        Custom Objects cannot be created via the API - they must be created on your Freshservice instance before they can be used.
    .EXAMPLE
        Get-FreshCustomObject
        Lists the available custom objects' id, title, description, updated_at & last_updated_by properties.
    .EXAMPLE
        Get-FreshCustomObject -CustomObjectID 1234567
        Will show more details about the custom object id 1234567, including id, name, title, description, (record) fields & meta details.
    .EXAMPLE
        Get-FreshCustomObject | Get-FreshCustomObject | Select-Object title,@{name='RecordCount';exp={$_.meta.total_records_count}}
        This will list all the custom objects, and show the number of records within each one.
    #>
    [CmdletBinding(DefaultParameterSetName='custom_objects')]
    param (
        # Custom object ID
        [parameter(Mandatory=$true,
            ParameterSetName='custom_object',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('custom_object_id')]
        [int64]$CustomObjectID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment             
    )

    process {
        $path = "objects"
        $field = $PSCmdlet.ParameterSetName

        if ($CustomObjectID)
        {
            $path += "/$CustomObjectID"
        }

        try {
            Invoke-FreshAPIGet -path $path -field $field -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')  | Select-Object *,@{name='custom_object_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function New-FreshCustomObjectRecord {
    <#
    .SYNOPSIS
        Creates a new record within a Fresh Custom Object
    .DESCRIPTION
        This creates a new record within a Fresh Custom Object.
    .NOTES
        Run (Get-FreshCustomObject -CustomObjectID 12345).fields to see the fields (required and otherwise), and where applicable, available choices.
    .EXAMPLE
        New-FreshCustomObjectRecord -CustomObjectID 123 -Data @{agent_group=5175566;approver=1232007392;category_dd1="Hardware";category_dd2="Computer";category_dd3="PC";item_name="1011125282";vendor_information="Apple-sales@apple.com"}
        This will add a new record to the Custom Object ID 123 with the details in the data.
    .EXAMPLE
        $NewRecordArray | New-FreshCustomObjectRecord -CustomObjectID 123 
        This will add all the records contained in $NewRecordArray to the Custom Object 123.
    #>
    [CmdletBinding()]
    param(
        # ID of the Custom object to which to add the record
        [parameter(Mandatory=$true)]
        [Alias('custom_object_id')]
        [int64]$CustomObjectID,

        # Record data contained in a hashtable
        [parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [hashtable]$Data,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    process {
        $body = @{
            data = $Data
        }

        $jsonBody = $body | ConvertTo-Json -Depth 10

        try {
            (Invoke-FreshAPIPost -path "objects/$CustomObjectID/records" -field "custom_object" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')).data | Select-Object *,@{name='record_id';exp={$_.bo_display_id}},@{name='custom_object_id';exp={$CustomObjectID}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }            
    }
    
}

function Get-FreshCustomObjectRecord {
    <#
    .SYNOPSIS
        This lists all the records present in a given Custom object.
    .DESCRIPTION
        This lists all the records present in a given Custom object.
    .NOTES
        The API supports queries and sorting, but as these can also be managed by PowerShell Where-Object and Sort-Object, I did not add query support into this function.
    .EXAMPLE
        Get-FreshCustomObject | Where-Object title -eq 'My Applications' | Get-FreshCustomObjectRecord
        This will list all records in the Custom Object with the title 'My Applications'
    #>
    [CmdletBinding()]
    param (
        # Custom object ID
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('custom_object_id')]
        [int64]$CustomObjectID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment          
    )
    
    process {
        # making page size the maximum value of 100
        $path = "objects/$CustomObjectID/records?page_size=100"
        $field = "records"

        try {
            (Invoke-FreshAPIGet -path $path -field $field -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') ).data | Select-Object *,@{name='record_id';exp={$_.bo_display_id}},@{name='custom_object_id';exp={$CustomObjectID}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }        
    }
}

function Update-FreshCustomObjectRecord {
    <#
    .SYNOPSIS
        Updates a record within a Fresh Custom Object
    .DESCRIPTION
        This updates a record within a Fresh Custom Object.
    .EXAMPLE
        Update-FreshCustomObjectRecord -CustomObjectID 123 -RecordID 2 -Data @{category_dd1="Hardware";category_dd2="Computer";category_dd3="Mac";item_name="1011125282"}
        This will update the record in the Custom Object ID 123 with the details in the data.
    .EXAMPLE
        Get-FreshCustomObjectRecord -CustomObjectID 123 | Where-Object Name -eq "My Application" | Update-FreshCustomObjectRecord -Data @{name='First Application'}
        This will update the record returned (where the name is 'My Application') and set the name to 'First Application'
    #>
    [CmdletBinding()]
    param(
        # ID of the Custom object to which to add the record
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('custom_object_id')]
        [int64]$CustomObjectID,

        # ID of the record to update
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('record_id')]
        $RecordID,

        # Record data contained in a hashtable
        [parameter(Mandatory=$true)]
        [hashtable]$Data,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    process {
        $body = @{
            data = $Data
        }

        $jsonBody = $body | ConvertTo-Json -Depth 10

        try {
            (Invoke-FreshAPIPut -path "objects/$CustomObjectID/records/$RecordID" -field "custom_object" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')).data | Select-Object *,@{name='record_id';exp={$_.bo_display_id}},@{name='custom_object_id';exp={$CustomObjectID}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }            
    }    
}

function Remove-FreshCustomObjectRecord {
    <#
    .SYNOPSIS
        Deletes a record from the Fresh Custom Object
    .DESCRIPTION
        Deletes a record from the Fresh Custom Object
    .NOTES
        Information or caveats about the function e.g. 'This function is not supported in Linux'
    .EXAMPLE
        Test-MyTestFunction -Verbose
        Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
    #>
    [CmdletBinding()]
    param (
        # ID of the Custom object to which to add the record
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('custom_object_id')]
        [int64]$CustomObjectID,

        # ID of the record to update
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('record_id')]
        $RecordID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            Invoke-FreshAPIDelete -path "objects/$CustomObjectID/records/$RecordID" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') 
        } catch {
            $_ | Convert-FreshError
        }              
    }
}