# Module for Fresh Department functions
# Les Newbigging 2022
#
# See https://api.freshservice.com/#departments

function New-FreshDepartment {
    <#
    .SYNOPSIS
        Create a new Department (or Company in MSP Mode) in Freshservice.
    .DESCRIPTION
        Create a new Department (or Company in MSP Mode) in Freshservice.
    .EXAMPLE
        New-FreshDepartment -Name "Researcg & Development" 
        Creates a new Fresh department called 'Research & Development'
    .EXAMPLE
        New-FreshDepartment -Name "IT Service Desk" -Description "The Service Desk for all IT enquiries." -HeadUserID (Get-FreshRequester -email servicedesk.supervisor@mycompany.com -IncludeAgents)
        Creates a new department in Fresh, with the head of department being the requester with the email of servicedesk.supervisor@mycompany.com
    .INPUTS
        None. Does not accept pipeline input.
    .OUTPUTS
        Object. A representation of the new department.
    #>    
    [CmdletBinding()]
    param(
        # Name of the department
        [parameter(Mandatory=$true)]
        [string]$Name,

        # Description about the department
        [string]$Description,

        # Unique identifier of the agent or requester who serves as the head of the department
        [Alias('head_user_id')]
        [int64]$HeadUserID,

        # Unique identifier of the agent or requester who serves as the prime user of the department
        [Alias('prime_user_id')]
        [int64]$PrimeUserID,

        # Email domains associated with the department
        [string[]]$Domains,

        # Custom fields that are associated with departments
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment   
    )

    $Body = @{
        name = $Name
        description = $Description
        head_user_id = $HeadUserID
        prime_user_id = $PrimeUserID
        domains = $Domains -join ','
        custom_fields = $CustomFields
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
        Invoke-FreshAPIPost -path "departments" -field "department" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')  | Select-Object *,@{name='department_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }    
}

function Get-FreshDepartment {
    <#
    .SYNOPSIS
        Retrieves a Fresh department or list of departments
    .DESCRIPTION
        Retrieves a list of departments, or a single department by department_id or name
    .INPUTS
        Object. Should include the department_id or department_ids property.
    .OUTPUTS
        Object[]. A representation of the Fresh department(s)
    .EXAMPLE
        Get-FreshDepartment
        Retrieves a list of all departments
    .EXAMPLE
        Get-FreshDepartment -Name "Finance"
        retrieves details for the finance department
    .EXAMPLE
        Get-FreshRequester -email fred.bloggs@bloggsblogs.org | Get-FreshDepartment | Select-Object department_id,name
        Will get ids and names of all the deparments Fred Bloggs is associated with.
    #>
    [CmdletBinding(DefaultParameterSetName='departments')]
    param (
        # Fresh ID for the department
        [parameter(Mandatory=$true,
            ParameterSetName='department',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('department_id','department_ids')]
        [int64[]]$DepartmentID,

        # Name of the department
        [parameter(Mandatory=$false,
            ParameterSetName='departments')]
        [string]$Name,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        if ($DepartmentID.count -gt 1)
        {
            foreach ($id in $DepartmentID)
            {
                # This will cater for multple ID values.
                Get-FreshDepartment -DepartmentID $id -System $System
            }
        } else {
            $path = "departments"
            $field = $PSCmdlet.ParameterSetName

            if ($DepartmentID)
            {
                $path += "/$DepartmentID"
            } else {
                $path += "?per_page=100"
                if ($Name)
                {
                    $path += '&query="name:' + "'$Name'" + '"'
                }
            }

            try {
                Invoke-FreshAPIGet -path $path -field $field -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')  | Select-Object *,@{name='department_id';exp={$_.id}},@{name='system';exp={$System}}
            } catch {
                $_ | Convert-FreshError
            }
        }
    }
}

function Update-FreshDepartment {
    <#
    .SYNOPSIS
        Update an existing Department (or Company in MSP Mode) in Freshservice.
    .DESCRIPTION
        Update an existing Department (or Company in MSP Mode) in Freshservice.
    .INPUTS
        Object. Must include department_id as a property
    .OUTPUTS
        Object. A representation of the updated department record.
    .EXAMPLE
        Update-FreshDepartment -DepartmentID 37 -Name "Help Desk"
        Changes the department name to 'Help Desk'
    #>
    [CmdletBinding()]
    param (
        # Fresh ID for the department
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('department_id')]
        [int64]$DepartmentID,

        # Name of the department
        [string]$Name,

        # Description about the department
        [string]$Description,

        # Unique identifier of the agent or requester who serves as the head of the department
        [Alias('head_user_id')]
        [int64]$HeadUserID,

        # Unique identifier of the agent or requester who serves as the prime user of the department
        [Alias('prime_user_id')]
        [int64]$PrimeUserID,

        # Email domains associated with the department
        [string[]]$Domains,

        # Custom fields that are associated with a Freshservice entity
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment              
    )

    process {

        $Body = @{
            name = $Name
            description = $Description
            head_user_id = $HeadUserID
            prime_user_id = $PrimeUserID
            domains = $Domains -join ','
            custom_fields = $CustomFields
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
            Invoke-FreshAPIPut -path "departments/$DepartmentID" -field "department" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')  | Select-Object *,@{name='department_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }        
    }
}

function Remove-FreshDepartment {
    <#
    .SYNOPSIS
        Delete the Department (or Company in MSP Mode) with the given ID from Freshservice.
    .DESCRIPTION
        Delete the Department (or Company in MSP Mode) with the given ID from Freshservice.
    .EXAMPLE
        Remove-FreshDepartment -DepartmentID 22
        Deletes the department with the id of 22 from Fresh
    .EXAMPLE
        Get-FreshDepartment | Where-Object Name -like 'IT*' | Remove-FreshDepartment
        Removes all departments whose names begin with 'IT'
    .INPUTS
        Object. Must have a property 'department_id'
    #>
    [CmdletBinding()]
    param (
        # Fresh ID for the department
        [parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true)]
        [Alias('department_id')]
        [int64]$DepartmentID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment            
    )
    
    process {
        try {
            Invoke-FreshAPIDelete -path "departments/$DepartmentID" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') 
        } catch {
            $_ | Convert-FreshError
        }             
    }
}

function Get-FreshDepartmentFields {
    <#
    .SYNOPSIS
        Retrieve the Department Fields (or Company Fields in MSP Mode) from Freshservice. 
    .DESCRIPTION
        Retrieve the Department Fields (or Company Fields in MSP Mode) from Freshservice. The fields will be returned in the sequence that they are displayed on the UI.
    .EXAMPLE
        Get-FreshDepartmentFields | Where-Object Mandatory -eq $true | Select-Object Name,Label
        Gets a list of mandatory fields, displaying their names and labels.
    .INPUTS
        None. This does not accept pipeline input.
    .OUTPUTS
        Object[]. An array of objects representing the department fields.
    #>
    [CmdletBinding()]
    param (
        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment          
    )
    
    try {
        Invoke-FreshAPIGet -path "department_fields" -field "department_fields" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }      
}