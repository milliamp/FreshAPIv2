# Module for Fresh Requester, Agent, and Group related functions.
# Les Newbigging 2022
# 
# See the following:
# Requesters: https://api.freshservice.com/#requesters
# Agents: https://api.freshservice.com/#agents
# (Agent) Roles: https://api.freshservice.com/#agent-roles
# Agent Groups: https://api.freshservice.com/#agent-groups
# requester Groups: https://api.freshservice.com/#requester-groups

# REQUESTER RELATED FUNCTIONS
function New-FreshRequester {
    <#
    .SYNOPSIS
        Creates a new Fresh Requester
    .DESCRIPTION
        Creates a new Fresh Requester.
    .PARAMETER System
        Which Fresh system/environment to use (Live or Sandbox).
    .EXAMPLE
        New-FreshRequester -email fred.flintstone@bedrock.org -FirstName Fred -LastName Flintstone 
        Creates a new Fresh Requester for Fred Flintstone
    .INPUTS
        None. No pipeline input taken.
    .OUTPUTS
        Object. A representation of the new requester.
    #> 
    [CmdletBinding()]
    param(
        # First name of the requester [MANDATORY]
        [parameter(Mandatory=$true)]
        [Alias('first_name')]
        [string]$FirstName,

        # Last name of the requester
        [Alias('last_name')]
        [string]$LastName,

        # Job title of the requester
        [Alias('job_title')]
        [string]$JobTitle,

        # Primary email address of the requester
        [parameter(Mandatory=$true,
            ParameterSetName='primary_email')]
        [Alias('primary_email','email')]
        [string]$PrimaryEmail,

        # Additional/secondary emails associated with the requester
        [Alias('secondary_emails')]
        [string[]]$SecondaryEmails,

        # Work phone number of the requester
        [parameter(Mandatory=$false,
            ParameterSetName='primary_email')]
        [parameter(Mandatory=$true,
            ParameterSetName='work_phone_number')]
        [Alias('work_phone_number')]
        [string]$WorkPhoneNumber,

        # Mobile phone number of the requester
        [parameter(Mandatory=$false,
            ParameterSetName='primary_email')]
        [parameter(Mandatory=$false,
            ParameterSetName='work_phone_number')]
        [parameter(Mandatory=$true,
            ParameterSetName='mobile_phone_number')]          
        [Alias('mobile_phone_number')]
        [string]$MobilePhoneNumber,

        # Unique IDs of the departments associated with the requester
        [Alias('department_ids')]
        [int64[]]$DepartmentIDs,

        # Set if the requester must be allowed to view tickets filed by other members of the department
        [Alias('can_see_all_tickets_from_associated_departments')]
        [switch]$CanSeeAllTicketsFromDepts,

        # User ID of the requester’s reporting manager
        [Alias('reporting_manager_id')]
        [int64]$ReportingManagerID,

        # Address of the requester        
        [string]$Address,

        # Time zone of the requester. For more information, see: https://support.freshservice.com/en/support/solutions/articles/232302-list-of-time-zones-supported-in-freshservice
        [Alias('time_zone')]
        [string]$TimeZone,

        # Time format for the requester.Possible values: 12h (12 hour format) / 24h (24 hour format)
        [ValidateSet('12h','24h')]
        [Alias('time_format')]
        [string]$TimeFormat,

        # Language used by the requester. The default language is “en” (English). Read more at https://support.freshservice.com/en/support/solutions/articles/232303-list-of-languages-supported-in-freshservice
        [string]$Language,
        
        # Unique ID of the location associated with the requester
        [Alias('location_id')]
        [int64]$LocationID,

        # Background information of the requester
        [Alias('background_information')]
        [string]$BackgroundInformation,

        # Key-value pair containing the names and values of the (custom) requester fields in a hashtable
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    $Body = @{
        first_name = $FirstName
        last_name = $LastName
        job_title = $JobTitle
        primary_email = $PrimaryEmail
        secondary_emails = $SecondaryEmails
        work_phone_number = $WorkPhoneNumber
        mobile_phone_number = $MobilePhoneNumber
        department_ids = $DepartmentIDs
        can_see_all_tickets_from_associated_departments = $CanSeeAllTicketsFromDepts.IsPresent
        reporting_manager_id = $ReportingManagerID
        address = $Address
        time_zone = $TimeZone
        time_format = $TimeFormat
        language = $Language
        location_id = $LocationID
        background_information = $BackgroundInformation
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
        Invoke-FreshAPIPost -path "requesters" -field "requester" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Get-FreshRequester {
    <#
    .SYNOPSIS
        Retrieves a list of Fresh requesters
    .DESCRIPTION
        Retrieves a list of requesters, or an individual requester, based on the parameters submitted.
    .EXAMPLE
        Get-FreshRequester
        Retrieves all Fresh requesters on the system.
    .EXAMPLE
        Get-FreshRequester -LastName Smith -IncludeAgents
        Gets a list of Fresh Requesters and Agents, whose surname is 'Smith'
    .OUTPUTS
        Object[]. An array of object that represent the requesters.
    #>
    [CmdletBinding(DefaultParameterSetName="requesters")]
    param(
        # Fresh Requester ID
        [parameter(Mandatory=$true,
                    position=0,
                    parametersetname="requester")]        
        [Alias('requester_id','id')]
        [int64]$RequesterID,

        # Email of requester
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]         
        [string]$Email,

        # Mobile phone number of user
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('mobile_phone_number','mobile')]
        [string]$MobilePhoneNumber,

        # Work phone number of user
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('work_phone_number','work_phone','work','work_number')]
        [string]$WorkPhoneNumber,

        # First name of the requester
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('first_name')]
        [string]$FirstName,

        # Last name of the requester
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('last_name')]
        [string]$LastName,

        # Concatenation of first_name and last_name with single space in-between fields
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [string]$Name,

        # Job title of the requester
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('job_title')]
        [string]$JobTitle,

        # (Primary) email address of the requester
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('primary_email')]
        [string]$PrimaryEmail,

        # ID of the department(s) assigned to the requester
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('department_id')]
        [int64]$DepartmentID,

        # ID of the reporting manager
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('reporting_manager_id')]
        [int64]$ReportingManagerID,

        # Time Zone (see list of time zones here: https://support.freshservice.com/en/support/solutions/articles/232302-list-of-time-zones-supported-in-freshservice)
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('time_zone')]
        [string]$TimeZone,

        # Language code (Eg. en, ja-JP) see article here: https://support.freshservice.com/en/support/solutions/articles/232303-list-of-languages-supported-in-freshservice
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [string]$Language,

        # ID of the location
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [Alias('location_id')]
        [int64]$LocationID,

        # Include Agents in the Requester results
        [parameter(Mandatory=$false,
                    parametersetname="requesters")]
        [switch]$IncludeAgents,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions

    $path = "requesters"
    $field = $PSCmdlet.ParameterSetName

    if ($field -eq 'requester')
    {
        # single id
        $path += "/$RequesterID"
    } else {
        # filters/all requesters
        $FilterParameters=@()

        # 100 is the maximum allowed page size, and will minimise the numer of calls made to the API, as the default is 30.
        $FilterParameters += "per_page=100" 

        if ($Email)
        {
            $FilterParameters += "email=$Email" 
        }

        if ($IncludeAgents)
        {
            $FilterParameters += "include_agents=true"
        }

        # if any query parameters are included, they'll be added here.
        $Queries = @()

        if ($FirstName)
        {
            $Queries += "first_name:'$FirstName'"
        }

        if ($LastName)
        {
            $Queries += "last_name:'$LastName'"
        }

        if ($Name)
        {
            $Queries += "name:'$Name'"
        }

        if ($JobTitle)
        {
            $Queries += "job_title:'$JobTitle'"
        }

        if ($PrimaryEmail)
        {
            $Queries += "primary_email:'$PrimaryEmail'"
        }

        if ($DepartmentID)
        {
            $Queries += "department_id:$DepartmentID"
        }

        if ($ReportingManagerID)
        {
            $Queries += "reporting_manager_id:$ReportingManagerID"
        }

        if ($TimeZone)
        {
            $Queries += "time_zone:'$TimeZone'"
        }

        if ($Language)
        {
            $Queries += "language:'$Language'"
        }

        if ($LocationID)
        {
            $Queries += "location_id:$LocationID"
        }

        # These could either be 'filters' or 'queries'. If there are existing queries, add to those; otherwise they'll be filters.
        if ($MobilePhoneNumber)
        {
            if ($Queries.count -eq 0)
            {
                # No other queries, so use as filter
                $FilterParameters += "mobile_phone_number=$MobilePhoneNumber"
            } else {
                # add to query list
                $Queries += "mobile_phone_number:'$MobilePhoneNumber'"
            }
            
        }

        if ($WorkPhoneNumber)
        {            
            if ($Queries.count -eq 0)
            {
                # No other queries, so use as filter
                $FilterParameters += "work_phone_number=$WorkPhoneNumber"
            } else {
                # add to query list
                $Queries += "work_phone_number:'$WorkPhoneNumber'"
            }            
        }

        # Build path string
        # Add queries to the filter
        if ($Queries.count -gt 0)
        {
            $FilterParameters += "query=" + ($Queries -join ' AND ')
        }

        $path += "?" + ($FilterParameters -join "&")
    }

    try {
        Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Update-FreshRequester {
    <#
    .SYNOPSIS
        Updates a Fresh Requester
    .DESCRIPTION
        Updates a Fresh Requester.
    .EXAMPLE
        Update-FreshRequester -RequesterID 5705 -TimeFormat "12h"  
        Updates the time format to 12 hour clock for the requester.
    .EXAMPLE
         Get-FreshRequester -LocationID 999 | Update-FreshRequester -LocationID 1000
         This will update all requesters at location id 999 to be at location ID 1000.
    .INPUTS
        Object. A representation of the requester. Must include requester_id
    .OUTPUTS
        Object. A representation of the requester.
    #> 
    [CmdletBinding()]
    param(
        # ID of Fresh Requester to update
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # First name of the requester
        [Alias('first_name')]
        [string]$FirstName,

        # Last name of the requester
        [Alias('last_name')]
        [string]$LastName,

        # Job title of the requester
        [Alias('job_title')]
        [string]$JobTitle,

        # Primary email address of the requester
        [Alias('primary_email','email')]
        [string]$PrimaryEmail,

        # Additional/secondary emails associated with the requester
        [Alias('secondary_emails')]
        [string[]]$SecondaryEmails,

        # Work phone number of the requester
        [Alias('work_phone_number')]
        [string]$WorkPhoneNumber,

        # Mobile phone number of the requester
        [Alias('mobile_phone_number')]
        [string]$MobilePhoneNumber,

        # Unique IDs of the departments associated with the requester
        [Alias('department_ids')]
        [int64[]]$DepartmentIDs,

        # Set if the requester must be allowed to view tickets filed by other members of the department
        [Alias('can_see_all_tickets_from_associated_departments')]
        [switch]$CanSeeAllTicketsFromDepts,

        # User ID of the requester’s reporting manager
        [Alias('reporting_manager_id')]
        [int64]$ReportingManagerID,

        # Address of the requester        
        [string]$Address,

        # Time zone of the requester. For more information, see: https://support.freshservice.com/en/support/solutions/articles/232302-list-of-time-zones-supported-in-freshservice
        [Alias('time_zone')]
        [string]$TimeZone,

        # Time format for the requester.Possible values: 12h (12 hour format) / 24h (24 hour format)
        [ValidateSet('12h','24h')]
        [Alias('time_format')]
        [string]$TimeFormat,

        # Language used by the requester. The default language is “en” (English). Read more at https://support.freshservice.com/en/support/solutions/articles/232303-list-of-languages-supported-in-freshservice
        [string]$Language,
        
        # Unique ID of the location associated with the requester
        [Alias('location_id')]
        [int64]$LocationID,

        # Background information of the requester
        [Alias('background_information')]
        [string]$BackgroundInformation,

        # Key-value pair containing the names and values of the (custom) requester fields in a hashtable
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    process {
        $Body = @{
            first_name = $FirstName
            last_name = $LastName
            job_title = $JobTitle
            primary_email = $PrimaryEmail
            secondary_emails = $SecondaryEmails
            work_phone_number = $WorkPhoneNumber
            mobile_phone_number = $MobilePhoneNumber
            department_ids = $DepartmentIDs
            can_see_all_tickets_from_associated_departments = $CanSeeAllTicketsFromDepts.IsPresent
            reporting_manager_id = $ReportingManagerID
            address = $Address
            time_zone = $TimeZone
            time_format = $TimeFormat
            language = $Language
            location_id = $LocationID
            background_information = $BackgroundInformation
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
            Invoke-FreshAPIPut -path "requesters/$RequesterID" -field "requester" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Remove-FreshRequester {
    <#
    .SYNOPSIS
        This function deactivates/forgets a requester
    .DESCRIPTION
        This function deactivates or forgets a user. Note that users from the sandbox cannot be forgotten.
    .EXAMPLE
        Remove-FreshRequester -RequesterID 876345934
        Deactivates the requester with ID 876345934
    .EXAMPLE
        Remove-FreshRequester -RequesterID 876345934 -Forget
        Forgets (deletes) the requester with ID 876345934        
    .OUTPUTS
        None.
    .INPUTS
        Object. This should have a property of requester_id at a minumum.
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Should this user be forgotten? This will remove the requester and any tickets raised by them.
        [switch]$Forget,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        if ($System -eq 'Sandbox' -and $Forget)
        {
            throw "Cannot forget Sandbox requesters."
        }
        $path = "requesters/$RequesterID"

        if ($Forget)
        {
            $path += "/forget"
        }
        try {
            Invoke-FreshAPIDelete -path $path -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

function Restore-FreshRequester {
    <#
    .SYNOPSIS
        This function reactivates a requester
    .DESCRIPTION
        This function eeactivates a user. Forgotten requesters cannot be reactivated.
    .EXAMPLE
        Restore-FreshRequester -RequesterID 876345934
        Reactivates the requester with ID 876345934
    .OUTPUTS
        Object. This is a representation of the requester.
    .INPUTS
        Object. This should have a property of requester_id at a minumum.
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            Invoke-FreshAPIPut -path "requesters/$RequesterID/reactivate" -field "requester" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

function Merge-FreshRequester {
    <#
    .SYNOPSIS
        This function merges secondary requesters into a primary requester.
    .DESCRIPTION
        This function merges secondary requesters into a primary requester.
    .EXAMPLE
        Merge-FreshRequester -RequesterID 111 -SecondaryRequesterID 222,333,444
        Merges requester IDs 222, 333 & 444 into 111.
    .OUTPUTS
        Object. A representation of the updated (primary) requester details.
    .INPUTS
        Object. This should have a property of requester_id at a minumum.
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Secondary Requester IDs to merge in to primary
        [parameter(Mandatory=$true)]
        [int64[]]$SecondaryRequesterID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        $path = "requesters/$RequesterID/merge?secondary_requesters=" + ($SecondaryRequesterID -join ',')

        try {
            Invoke-FreshAPIPut -path $path -field "requester" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

# AGENT RELATED FUNCTIONS
function New-FreshAgent {
    <#
    .SYNOPSIS
        Creates a new Fresh Agent
    .DESCRIPTION
        This operation allows you to create a new agent in Freshservice.
    .EXAMPLE
        New-FreshAgent -email fred.flintstone@bedrock.org -FirstName Fred -LastName Flintstone 
        Creates a new Fresh Agent for Fred Flintstone
    .INPUTS
        None. No pipeline input taken.
    .OUTPUTS
        Object. A representation of the new agent.
    .NOTES
        ROLES:
        Each individual role is a hash in the roles array that contains the attributes. 
        + role_id:          Unique ID of the role assigned
        + assignment_scope: The scope in which the agent can use the permissions granted by this role. Possible values include entire_helpdesk (all plans)
                              member_groups (all plans; in the Pro and Enterprise plans, this also includes groups that the agent is an observer of)
                              specified_groups (Pro and Enterprise only), and assigned_items (all plans)
        + groups:           Unique IDs of Groups in which the permissions granted by the role applies. Mandatory only when the assignment_scope is specified_groups, and should be ignored otherwise.
    #> 
    [CmdletBinding()]
    param(
        # First name of the requester [MANDATORY]
        [parameter(Mandatory=$true)]
        [Alias('first_name')]
        [string]$FirstName,

        # Last name of the requester
        [Alias('last_name')]
        [string]$LastName,

        # Set this if the agent is an 'occasional' agent (not full-time agent)
        [switch]$Occasional,

        # Job title of the requester
        [Alias('job_title')]
        [string]$JobTitle,

        # Primary email address of the requester
        [parameter(Mandatory=$true,
            ParameterSetName='primary_email')]
        [Alias('primary_email','email')]
        [string]$Email,

        # Work phone number of the requester
        [parameter(Mandatory=$false,
            ParameterSetName='primary_email')]
        [parameter(Mandatory=$true,
            ParameterSetName='work_phone_number')]
        [Alias('work_phone_number')]
        [string]$WorkPhoneNumber,

        # Mobile phone number of the requester
        [parameter(Mandatory=$false,
            ParameterSetName='primary_email')]
        [parameter(Mandatory=$false,
            ParameterSetName='work_phone_number')]
        [parameter(Mandatory=$true,
            ParameterSetName='mobile_phone_number')]          
        [Alias('mobile_phone_number')]
        [string]$MobilePhoneNumber,

        # Unique IDs of the departments associated with the requester
        [Alias('department_ids')]
        [int64[]]$DepartmentIDs,

        # Set if the requester must be allowed to view tickets filed by other members of the department
        [Alias('can_see_all_tickets_from_associated_departments')]
        [switch]$CanSeeAllTicketsFromDepts,

        # User ID of the requester’s reporting manager
        [Alias('reporting_manager_id')]
        [int64]$ReportingManagerID,

        # Address of the requester        
        [string]$Address,

        # Time zone of the requester. For more information, see: https://support.freshservice.com/en/support/solutions/articles/232302-list-of-time-zones-supported-in-freshservice
        [Alias('time_zone')]
        [string]$TimeZone,

        # Time format for the requester.Possible values: 12h (12 hour format) / 24h (24 hour format)
        [ValidateSet('12h','24h')]
        [Alias('time_format')]
        [string]$TimeFormat,

        # Language used by the requester. The default language is “en” (English). Read more at https://support.freshservice.com/en/support/solutions/articles/232303-list-of-languages-supported-in-freshservice
        [string]$Language,
        
        # Unique ID of the location associated with the requester
        [Alias('location_id')]
        [int64]$LocationID,

        # Background information of the requester
        [Alias('background_information')]
        [string]$BackgroundInformation,

        # Unique ID of the level of the agent in the Arcade. Possible values: 1 (Beginner), 2 (Intermediate), 3 (Professional), 4 (Expert), 5 (Master), 6 (Guru)
        [ValidateRange(1,6)]
        [Alias('scoreboard_level_id')]
        [int]$ScoreboardLevel,

        # Unique IDs of the groups that the agent is a member of. The response value for this field would only contain the list of groups that the agent is an approved member of. The member_of_pending_approval read-only attribute in the response will include the list of groups for which the agent’s member access is pending approval by a group leader.
        [Alias('member_of','group_id')]
        [int64[]]$MemberOf,

        # Unique IDs of the groups that the agent is an observer of. The response value for this field would only contain the list of groups that the agent is an approved observer of. The observer_of_pending_approval read-only attribute in the response will include the list of groups for which the agent’s observer access is pending approval by a group leader.
        [Alias('observer_of')]
        [int64[]]$ObserverOf,

        # See ROLES under NOTES.
        [Alias('role_id','role_ids')]
        [hashtable[]]$Role,

        # Signature of the agent in HTML format.
        [string]$Signature,

        # Key-value pair containing the names and values of the (custom) requester fields in a hashtable
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    $Body = @{
        first_name = $FirstName
        last_name = $LastName
        occasional = $Occasional.IsPresent
        job_title = $JobTitle
        email = $PrimaryEmail
        work_phone_number = $WorkPhoneNumber
        mobile_phone_number = $MobilePhoneNumber
        department_ids = $DepartmentIDs
        can_see_all_tickets_from_associated_departments = $CanSeeAllTicketsFromDepts.IsPresent
        reporting_manager_id = $ReportingManagerID
        address = $Address
        time_zone = $TimeZone
        time_format = $TimeFormat
        language = $Language
        location_id = $LocationID
        background_information = $BackgroundInformation
        scoreboard_level_id = $ScoreboardLevel
        member_of = $MemberOf
        observer_of = $ObserverOf
        roles = $Role
        signature = $Signature
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
        Invoke-FreshAPIPost -path "agents" -field "agent" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='agent_id';exp={$_.id}},@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Get-FreshAgent {
    <#
    .SYNOPSIS
        Retrieves a list of Fresh Agents
    .DESCRIPTION
        Retrieves a list of Agents, or an individual Agent, based on the parameters submitted.
    .EXAMPLE
        Get-FreshAgent
        Retrieves all Fresh requesters on the system.
    .EXAMPLE
        Get-FreshAgent -LastName Smith
        Gets a list of Fresh Agents, whose surname is 'Smith'
    .INPUTS
        Object. Must contain an agent_id field.
    .OUTPUTS
        Object[]. An array of objects representing the agents.
    #>
    [CmdletBinding(DefaultParameterSetName="agents")]
    param(
        # Search by ID of Fresh Agent required
        [parameter(Mandatory=$true,
                    position=0,
                    parametersetname="agent")]        
        [Alias('agent_id')]
        [int64]$AgentID,

        # Email address of the agent.
        [parameter(Mandatory=$false,
                    parametersetname="agents")]         
        [string]$Email,

        # Mobile phone number of agent
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('mobile_phone_number','mobile')]
        [string]$MobilePhoneNumber,

        # Work phone number of agent
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('work_phone_number','work_phone','work','work_number')]
        [string]$WorkPhoneNumber,

        # Agent active state
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [ValidateSet('Yes','No')]
        [string]$Active,

        # Agent 'State' (i.e. fulltime or occasional)
        [parameter(Mandatory=$false,
                    parametersetname="agents")]        
        [ValidateSet('fulltime','occasional')]
        $State,

        # First name of the agent
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('first_name')]
        [string]$FirstName,

        # Last name of the agent
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('last_name')]
        [string]$LastName,

        # Concatenation of first_name and last_name with single space in-between fields
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [string]$Name,

        # Job title of the agent
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('job_title')]
        [string]$JobTitle,

        # ID of the department(s) assigned to the requester
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('department_id')]
        [int64]$DepartmentID,

        # ID of the reporting manager
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('reporting_manager_id')]
        [int64]$ReportingManagerID,

        # Time Zone (see list of time zones here: https://support.freshservice.com/en/support/solutions/articles/232302-list-of-time-zones-supported-in-freshservice)
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('time_zone')]
        [string]$TimeZone,

        # Language code (Eg. en, ja-JP) see article here: https://support.freshservice.com/en/support/solutions/articles/232303-list-of-languages-supported-in-freshservice
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [string]$Language,

        # ID of the location
        [parameter(Mandatory=$false,
                    parametersetname="agents")]
        [Alias('location_id')]
        [int64]$LocationID,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions

    $path = "agents"
    $field = $PSCmdlet.ParameterSetName

    if ($field -eq 'agent')
    {
        # single id
        $path += "/$AgentID"
    } else {
        # filters/all requesters
        $FilterParameters=@()

        # 100 is the maximum allowed page size, and will minimise the numer of calls made to the API, as the default is 30.
        $FilterParameters += "per_page=100" 

        if ($Email)
        {
            $FilterParameters += "email=$Email" 
        }

        if ($Active)
        {
            $ActiveResult = "$($Active -eq 'Yes')".ToLower()
            $FilterParameters += "active=$ActiveResult" 
        }   
        
        if ($State)
        {
            $FilterParameters += "state=$($State.ToLower())" 
        }        

        # if any query parameters are included, they'll be added here.
        $Queries = @()

        if ($FirstName)
        {
            $Queries += "first_name:'$FirstName'"
        }

        if ($LastName)
        {
            $Queries += "last_name:'$LastName'"
        }

        if ($Name)
        {
            $Queries += "name:'$Name'"
        }

        if ($JobTitle)
        {
            $Queries += "job_title:'$JobTitle'"
        }

        if ($DepartmentID)
        {
            $Queries += "department_id:$DepartmentID"
        }

        if ($ReportingManagerID)
        {
            $Queries += "reporting_manager_id:$ReportingManagerID"
        }

        if ($TimeZone)
        {
            $Queries += "time_zone:'$TimeZone'"
        }

        if ($Language)
        {
            $Queries += "language:'$Language'"
        }

        if ($LocationID)
        {
            $Queries += "location_id:$LocationID"
        }

        # These could either be 'filters' or 'queries'. If there are existing queries, add to those; otherwise they'll be filters.
        if ($MobilePhoneNumber)
        {
            if ($Queries.count -eq 0)
            {
                # No other queries, so use as filter
                $FilterParameters += "mobile_phone_number=$MobilePhoneNumber"
            } else {
                # add to query list
                $Queries += "mobile_phone_number:'$MobilePhoneNumber'"
            }
            
        }

        if ($WorkPhoneNumber)
        {            
            if ($Queries.count -eq 0)
            {
                # No other queries, so use as filter
                $FilterParameters += "work_phone_number=$WorkPhoneNumber"
            } else {
                # add to query list
                $Queries += "work_phone_number:'$WorkPhoneNumber'"
            }            
        }

        # Build path string
        # Add queries to the filter
        if ($Queries.count -gt 0)
        {
            $FilterParameters += "query=" + ($Queries -join ' AND ')
        }

        $path += "?" + ($FilterParameters -join "&")
    }

    try {
        Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='agent_id';exp={$_.id}},@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Update-FreshAgent {
    <#
    .SYNOPSIS
        Updates a Fresh Agent
    .DESCRIPTION
        Updates a Fresh Agent.
    .EXAMPLE
        Update-FreshAgent -AgentID 5705 -TimeFormat "24h"  
        Updates the time format to 12 hour clock for the agent.
    .EXAMPLE
         Get-FreshAgent -LocationID 999 | Update-FreshAgent -LocationID 1000
         This will update all agents at location id 999 to be at location ID 1000.
    .INPUTS
        Object. A representation of the agent.
    .OUTPUTS
        Object. A representation of the (updated) agent.
    #> 
    [CmdletBinding()]
    param(
        # ID of Fresh Requester to update
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('agent_id')]
        [int64]$AgentID,

        # Set if the agent is an occasional agent
        [switch]$Occasional,

        # Email address of the agent
        [string]$Email,

        # Work phone number of the agent
        [Alias('work_phone_number')]
        [string]$WorkPhoneNumber,

        # Mobile phone number of the agent
        [Alias('mobile_phone_number')]
        [string]$MobilePhoneNumber,

        # Unique IDs of the departments associated with the agent
        [Alias('department_ids')]
        [int64[]]$DepartmentIDs,

        # Set if the requester must be allowed to view tickets filed by other members of the department
        [Alias('can_see_all_tickets_from_associated_departments')]
        [switch]$CanSeeAllTicketsFromDepts,

        # User ID of the requester’s reporting manager
        [Alias('reporting_manager_id')]
        [int64]$ReportingManagerID,

        # Address of the requester        
        [string]$Address,

        # Time zone of the requester. For more information, see: https://support.freshservice.com/en/support/solutions/articles/232302-list-of-time-zones-supported-in-freshservice
        [Alias('time_zone')]
        [string]$TimeZone,

        # Time format for the requester.Possible values: 12h (12 hour format) / 24h (24 hour format)
        [ValidateSet('12h','24h')]
        [Alias('time_format')]
        [string]$TimeFormat,

        # Language used by the requester. The default language is “en” (English). Read more at https://support.freshservice.com/en/support/solutions/articles/232303-list-of-languages-supported-in-freshservice
        [string]$Language,
        
        # Unique ID of the location associated with the requester
        [Alias('location_id')]
        [int64]$LocationID,

        # Background information of the requester
        [Alias('background_information')]
        [string]$BackgroundInformation,

        # Unique ID of the level of the agent in the Arcade. Possible values: 1 (Beginner), 2 (Intermediate), 3 (Professional), 4 (Expert), 5 (Master), 6 (Guru)
        [ValidateRange(1,6)]
        [Alias('scoreboard_level_id')]
        [int]$ScoreboardLevel,

        # Unique IDs of the groups that the agent is a member of. The response value for this field would only contain the list of groups that the agent is an approved member of. The member_of_pending_approval read-only attribute in the response will include the list of groups for which the agent’s member access is pending approval by a group leader.
        [Alias('member_of','group_id')]
        [int64[]]$MemberOf,

        # Unique IDs of the groups that the agent is an observer of. The response value for this field would only contain the list of groups that the agent is an approved observer of. The observer_of_pending_approval read-only attribute in the response will include the list of groups for which the agent’s observer access is pending approval by a group leader.
        [Alias('observer_of')]
        [int64[]]$ObserverOf,

        # See ROLES under NOTES for New-FreshAgent.
        [Alias('role_id','role_ids')]
        [hashtable[]]$Role,

        # Signature of the agent in HTML format.
        [string]$Signature,
        
        # Key-value pair containing the names and values of the (custom) requester fields in a hashtable
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    process {
        $Body = @{
            occasional = $Occasional.IsPresent
            email = $PrimaryEmail
            work_phone_number = $WorkPhoneNumber
            mobile_phone_number = $MobilePhoneNumber
            department_ids = $DepartmentIDs
            can_see_all_tickets_from_associated_departments = $CanSeeAllTicketsFromDepts.IsPresent
            reporting_manager_id = $ReportingManagerID
            address = $Address
            time_zone = $TimeZone
            time_format = $TimeFormat
            language = $Language
            location_id = $LocationID
            background_information = $BackgroundInformation
            scoreboard_level_id = $ScoreboardLevel
            member_of = $MemberOf
            observer_of = $ObserverOf
            roles = $Role
            signature = $Signature
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
            Invoke-FreshAPIPut -path "agents/$AgentID" -field "agent" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='agent_id';exp={$_.id}},@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Remove-FreshAgent {
    <#
    .SYNOPSIS
        This function deactivates an agent
    .DESCRIPTION
        This function deactivates an agent.
    .EXAMPLE
        Remove-FreshAgent -AgentID 876345934
        Deactivates the agent with ID 876345934
    .OUTPUTS
        None.
    .INPUTS
        Object. This should have a property of agent_id at a minumum.
    #>
    [CmdletBinding()]
    param(
        # Fresh agent ID to delete
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('agent_id')]
        [int64]$AgentID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            (Invoke-FreshAPIDelete -path "agents/$AgentID" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')).agent
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

function Restore-FreshAgent {
    <#
    .SYNOPSIS
        This function reactivates an agent
    .DESCRIPTION
        This function reactivates an agent
    .EXAMPLE
        Restore-FreshAgent -AgentID 876345934
        Reactivates the agent with ID 876345934
    .OUTPUTS
        Object. This is a representation of the requester.
    .INPUTS
        Object. This should have a property of requester_id at a minumum.
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('agent_id')]
        [int64]$AgentID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            Invoke-FreshAPIPut -path "agents/$AgentID/reactivate" -field "agent" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='agent_id';exp={$_.id}},@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

# CONVERSION FUNCTIONS
function ConvertTo-FreshAgent {
    <#
    .SYNOPSIS
        This function converts a requester to an agent.
    .DESCRIPTION
        This function converts a requester to an agent.
    .EXAMPLE
        ConvertTo-FreshAgent -RequesterID 280280
        Converts Request 280280 to an agent.
    .OUTPUTS
        Object. A representation of the (now converted) Agent details.
    .INPUTS
        Object. This should have a property of requester_id at a minumum.
    #>
    [CmdletBinding()]
    param(
        # Fresh requester ID to convert to agent
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            Invoke-FreshAPIPut -path "requesters/$RequesterID/convert_to_agent" -field "agent" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requester_id';exp={$_.id}},@{name='agent_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

function ConvertTo-FreshRequester {
    <#
    .SYNOPSIS
        This function converts an agent to a requester.
    .DESCRIPTION
        This function converts an agent to a requester.
    .EXAMPLE
        ConvertTo-FreshAgent -AgentID 280280
        Converts agent 280280 to a requester.
    .OUTPUTS
        Object. A representation of the (now converted) requester details.
    .INPUTS
        Object. This should have a property of agent_id at a minumum.
    #>
    [CmdletBinding()]
    param(
        # Fresh agent ID to convert
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('agent_id')]
        [int64]$AgentID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            Invoke-FreshAPIPut -path "agents/$RequesterID/convert_to_requester" -field "requester" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

# OTHER AGENT FUNCTIONS
function Get-FreshAgentField {
    <#
    .SYNOPSIS
        Retrieves a list of agent fields
    .DESCRIPTION
        Retrieves a list of agent fields
    .INPUTS
        None. This does not accept pipeline input.
    .OUTPUTS
        Object[]. An array of objects representing the fields included in an agent record.
    #>
    [CmdletBinding()]
    param (
        # Fresh system/environment to query
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment            
    )

    try {
        Invoke-FreshAPIGet -path "agent_fields" -field "agent_fields" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')
    } catch {
        $_ | Convert-FreshError
    }      

}

function Get-FreshAgentRole {
    <#
    .SYNOPSIS
        Retrieves agent roles
    .DESCRIPTION
        Gets a list of agent roles, or a single role if specified.
    .INPUTS
        Object. Must include the role_id or role_ids property.
    .OUTPUTS
        Object[]. An array of objects representing the roles.
    .EXAMPLE
        Get-FreshAgentRole | Select-Object role_id,name,description,scopes
        role_id name                  description                                                                                                                scopes
        ------- ----                  -----------                                                                                                                ------
            90 Account Admin         Has complete control over the help desk including access to Account or Billing related information, and receives Invoices. @{ticket=; problem=; change=; release=; asset=; solution=; ... 
            91 Admin                 Can configure all features through the Admin tab, but is restricted from viewing Account or Billing related information.   @{ticket=; problem=; change=; release=; asset=; solution=; ...
            92 SD Supervisor         Can perform all agent related activities and access reports, but cannot access or change configurations in the Admin tab.  @{ticket=; problem=; asset=; solution=; contract=}
            93 SD Agent              Can log, view, reply, update and resolve tickets and manage contacts.                                                      @{ticket=; problem=; change=; asset=; solution=; contract=}    
    .EXAMPLE
        Get-FreshAgent -Email winston.smith@bigbrother.gov | Get-FreshAgentRole | Select-Object Name
        Displays a list of roles held by Winston Smith
    .EXAMPLE
        Get-FreshAgentRole -RoleID 95
        id          : 95
        name        : Change Manager
        description : Can perform all agent related activities, view and create problem release and have full access for change module
        default     : True
        scopes      : @{ticket=; problem=; change=; release=; asset=; contract=}
        created_at  : 2020-12-25T14:00:00Z
        updated_at  : 2020-12-25T14:00:00Z
        role_id     : 95
        system      : live
    #>
    [CmdletBinding(DefaultParameterSetName='roles')]
    param (
        # Retrieve an individual role details
        [parameter(Mandatory=$true,
            ParameterSetName='role',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('role_id','role_ids')]
        $RoleID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment             
    )

    process {
        $field = $PSCmdlet.ParameterSetName
        $path = "roles/"
        
        if ($RoleID)
        {
            $path += "$RoleID"
        }

        try {
            Invoke-FreshAPIGet -path $Path -field $field -system $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='role_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }  
    }
}

# AGENT GROUP FUNCTIONS
function New-FreshAgentGroup {
    <#
    .SYNOPSIS
        Creates a new Fresh Agent group
    .DESCRIPTION
        This operation allows you to create a new agent group in Freshservice.
    .NOTES
        When creating groups with approval required only the leaders will be added to the groups immediately while the members and observers would get added after their membership approval is approved.
    .EXAMPLE
        New-FreshAgentGroup -Name "Office Management"
        Creates a new agent group called "Office Management"
    .INPUTS
        None. This does not accept pipeline input.
    .OUTPUTS
        Object. A representation of the new agent group.
    #> 
    [CmdletBinding()]
    param(
        # Name of the group.
        [parameter(Mandatory=$true)]
        [string]$Name,

        # Description of the group.
        [string]$Description,

        # The time after which an escalation email is sent if a ticket in the group remains unassigned. The accepted values are “30m” for 30 minutes, “1h” for 1 hour, “2h” for 2 hours“, “4h” for 4 hours, “8h” for 8 hours, “12h” for 12 hours, “1d” for 1 day, “2d” for 2 days, and “3d” for 3 days. Default is 30m.
        [ValidateSet('30m','1h','2h','4h','8h','12h','1d','2d','3d')]
        [Alias('unassigned_for')]
        [string]$UnassignedFor,

        # Unique ID of the business hours configuration associated with the group.
        [Alias('business_hours_id')]
        [int]$BusinessHoursID,

        # The Unique ID of the user to whom an escalation email is sent if a ticket in this group is unassigned. To create/update a group with an escalate_to value of ‘none’, please set the value of this parameter to ‘null’.
        [Alias('escalate_to')]
        [int64]$EscalateToID,

        # A comma separated array of user IDs of agents who are members of this group. The response value for this field would only contain the list of approved members. The members_pending_approval read-only attribute in the response will include the list of members whose approval is pending.
        [int64[]]$Members,

        # A comma separated array of user IDs of agents who are observers of this group. The response value for this field would only contain the list of approved observers. The observers_pending_approval read-only attribute in the response will include the list of observers whose approval is pending. This attribute is only applicable for accounts which have the “Access Controls Pro” feature enabled.
        [int64[]]$Observers,

        # Signifies whether a group is marked as restricted. This attribute won't be supported if the "Access Controls Pro" feature is unavailable for the account.
        [switch]$Restricted,

        # A comma separated array of user IDs of agents who are leaders of this group. The response value for this field would only contain the list of approved leaders. The leaders_pending_approval read-only attribute in the response will include the list of leaders whose approval is pending. This attribute is only applicable for accounts which have the “Access Controls Pro” feature enabled.
        [int64[]]$Leaders,

        # Signifies whether the restricted group requires approvals for membership changes. This attribute is only applicable for accounts which have the “Access Controls Pro” feature enabled.
        [Alias('approval_required')]
        [switch]$ApprovalRequired,

        # Describes the automatic ticket assignment type. Will not be supported if the "Round Robin" feature is disabled for the account.
        [Alias('auto_ticket_assign')]
        [switch]$AutoTicketAssign,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    $Body = @{
        name = $Name
        description = $Description
        unassigned_for = $UnassignedFor
        business_hours_id = $BusinessHoursID
        escalate_to = $EscalateToID
        members = $Members -join ','
        observers = $Observers -join ','
        restricted = $Restricted.IsPresent
        leaders = $Leaders -join ','
        approval_required = $ApprovalRequired.IsPresent
        auto_ticket_assign = $AutoTicketAssign.IsPresent
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
        Invoke-FreshAPIPost -path "groups" -field "group" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='group_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Get-FreshAgentGroup {
    <#
    .SYNOPSIS
        Retrieves a list of Fresh Agent groups
    .DESCRIPTION
        Retrieves a list of Agent groups, or an individual Agent group.
    .EXAMPLE
        Get-FreshAgentGroup
        Lists all agent groups.
    .EXAMPLE
        Get-FreshAgent -email evie@vendetta.org | Get-FreshAgentGroup | Select-Object Name,Description
        Returns the names and descriptions of the groups that Evie is a member of.
    .INPUTS
        Object. Requires either the group_id or group_ids property.
    .OUTPUTS
        Object[]. An array of objects representing the agent groups.
    #>
    [CmdletBinding(DefaultParameterSetName="groups")]
    param(
        # ID of group to view
        [parameter(Mandatory=$true,
            ParameterSetName="group",
            ValueFromPipelineByPropertyName=$true)]
        [Alias('group_id','group_ids')]
        $GroupID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )
    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions
    }

    process {
        $path = "groups"
        $field = $PSCmdlet.ParameterSetName

        if ($GroupID)
        {
            $path += "/$GroupID"
        } else {
            $path += "?per_page=100"
        }

        try {
            Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='group_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Update-FreshAgentGroup {
    <#
    .SYNOPSIS
        Updates a Fresh Agent group
    .DESCRIPTION
        This operation allows you to modify an agent group in Freshservice.
    .NOTES
        To delete all the agents associated with a group, update the group with "members"= [ ] (empty array)
        Appending the pending users from "members_pending_approval", "observers_pending_approval", "leaders_pending_approval" to "members", "observers", "leaders" keys respectively will retain the pending approval request for these users. Not doing so would withdraw their membership approval request.
        When a restricted group that required approvals is either marked unrestricted or approval required is set to false the pending membership request will get auto approved.
    .INPUTS
        Object. A representation of the agent group to modify - requires the group_id property.
    .OUTPUTS
        Object. A representation of the updated group.
    .EXAMPLE
        Get-FreshAgentGroup | Where-Object name -eq 'IT Service Desk' | Update-FreshAgentGroup -Name 'Help Desk' 
        Changes the name of the agent group from 'IT Service Desk' to 'Help Desk' 
    #> 
    [CmdletBinding()]
    param(
        # ID of the Fresh Agent group
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('group_id')]
        [int64]$GroupID,

        # Name of the group.
        [string]$Name,

        # Description of the group.
        [string]$Description,

        # The time after which an escalation email is sent if a ticket in the group remains unassigned. The accepted values are “30m” for 30 minutes, “1h” for 1 hour, “2h” for 2 hours“, “4h” for 4 hours, “8h” for 8 hours, “12h” for 12 hours, “1d” for 1 day, “2d” for 2 days, and “3d” for 3 days. Default is 30m.
        [ValidateSet('30m','1h','2h','4h','8h','12h','1d','2d','3d')]
        [Alias('unassigned_for')]
        [string]$UnassignedFor,

        # Unique ID of the business hours configuration associated with the group.
        [Alias('business_hours_id')]
        [int]$BusinessHoursID,

        # The Unique ID of the user to whom an escalation email is sent if a ticket in this group is unassigned. To create/update a group with an escalate_to value of ‘none’, please set the value of this parameter to ‘null’.
        [Alias('escalate_to')]
        [int64]$EscalateToID,

        # A comma separated array of user IDs of agents who are members of this group. If the group is restricted and approvals-enabled, the input value for this field should also include the user IDs of agents whose member access to the group is pending approval by a group leader. The response value for this field would only contain the list of approved members. The members_pending_approval read-only attribute in the response will include the list of members whose approval is pending.
        [int64[]]$Members,

        # A comma separated array of user IDs of agents who are observers of this group. If the group is restricted and approvals-enabled, the input value for this field should also include the user IDs of agents whose observer access to the group is pending approval by a group leader. The response value for this field would only contain the list of approved observers. The observers_pending_approval read-only attribute in the response will include the list of observers whose approval is pending. This attribute is only applicable for accounts which have the “Access Controls Pro” feature enabled.
        [int64[]]$Observers,

        # Signifies whether a group is marked as restricted. This attribute won't be supported if the "Access Controls Pro" feature is unavailable for the account.
        [switch]$Restricted,

        # A comma separated array of user IDs of agents who are leaders of this group. If the group is restricted and approvals-enabled, the input value for this field should also include the user IDs of agents whose leader access to the group is pending approval by another group leader. The response value for this field would only contain the list of approved leaders. The leaders_pending_approval read-only attribute in the response will include the list of leaders whose approval is pending. This attribute is only applicable for accounts which have the “Access Controls Pro” feature enabled.
        [int64[]]$Leaders,

        # Signifies whether the restricted group requires approvals for membership changes. This attribute is only applicable for accounts which have the “Access Controls Pro” feature enabled.
        [Alias('approval_required')]
        [switch]$ApprovalRequired,

        # Describes the automatic ticket assignment type. Will not be supported if the "Round Robin" feature is disabled for the account.
        [Alias('auto_ticket_assign')]
        [switch]$AutoTicketAssign,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    $Body = @{
        name = $Name
        description = $Description
        unassigned_for = $UnassignedFor
        business_hours_id = $BusinessHoursID
        escalate_to = $EscalateToID
        members = $Members -join ','
        observers = $Observers -join ','
        restricted = $Restricted.IsPresent
        leaders = $Leaders -join ','
        approval_required = $ApprovalRequired.IsPresent
        auto_ticket_assign = $AutoTicketAssign.IsPresent
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
        Invoke-FreshAPIPut -path "groups/$GroupID" -field "group" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='group_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Remove-FreshAgentGroup {
    <#
    .SYNOPSIS
        This function deletes an agent group
    .DESCRIPTION
        This function deletes an agent group.
    .INPUTS
        Object. A respresentation of the agent group to delete, requires the group_id property.
    .OUTPUTS
        None.
    .EXAMPLE
        Remove-FreshAgentGroup -GroupID 123
        Deletes the group with ID 123.
    #>
    [CmdletBinding()]
    param(
        # Fresh agent group ID to delete
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('group_id')]
        [int64]$GroupID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            (Invoke-FreshAPIDelete -path "groups/$GroupID" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')).group
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

# REQUESTER GROUP FUNCTIONS
function New-FreshRequesterGroup {
    <#
    .SYNOPSIS
        Creates a new Fresh requester group
    .DESCRIPTION
        This operation allows you to create a new requester group in Freshservice.
    .NOTES
        Only manual groups can be created or updated via the APIs
    .INPUTS
        None. This does not accept pipeline input.
    .OUTPUTS
        Object. A representation of the requester group.
    .EXAMPLE
        New-FreshRequesterGroup -Name "Directors" -Description "All the company directors"
        Creates a requester group called 'Directors'
    #> 
    [CmdletBinding()]
    param(
        # Name of the group.
        [parameter(Mandatory=$true)]
        [string]$Name,

        # Description of the group.
        [string]$Description,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    $Body = @{
        name = $Name
        description = $Description
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
        Invoke-FreshAPIPost -path "requester_groups" -field "requester_group" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requester_group_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Get-FreshRequesterGroup {
    <#
    .SYNOPSIS
        Retrieves a list of Fresh requester groups
    .DESCRIPTION
        Retrieves a list of requester groups, or an individual requester group.
    .NOTES
        Groups of both types (“manual”/”rule_based”) can be viewed
    .INPUTS
        Object. Requires the requester_group_id property.
    .OUTPUTS
        Object[]. An array of objects representing the requester groups.
    .EXAMPLE
        Get-FreshRequesterGroup
        Retrieves all Fresh requester groups
    #>
    [CmdletBinding(DefaultParameterSetName="requester_groups")]
    param(
        # ID of group to view
        [parameter(Mandatory=$true,
            ParameterSetName="requester_group",
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_group_id')]
        $RequesterGroupID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )
    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions
    }

    process {
        $path = "requester_groups"
        $field = $PSCmdlet.ParameterSetName

        if ($RequesterGroupID)
        {
            $path += "/$RequesterGroupID"
        } else {
            $path += "?per_page=100"
        }

        try {
            Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='requester_group_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Update-FreshRequesterGroup {
    <#
    .SYNOPSIS
        Updates a Fresh requester group
    .DESCRIPTION
        This operation allows you to modify a requester group in Freshservice.
    .NOTES
        Only manual groups can be created or updated via the APIs
    .INPUTS
        Object. A representation of the requester group to update. Must include the requester_group_id property.
    .OUTPUTS
        Object. A representation of the modified requester group.
    .EXAMPLE
        Update-FreshRequesterGroup -RequesterGroupID 777 -Name "Unix Admins" -Description "Requesters who are Unix Administrators"
        Updates the name and description of requester_group_id 777
    #> 
    [CmdletBinding()]
    param(
        # ID of Fresh requester group to update
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_group_id')]
        $RequesterGroupID,

        # Name of the group.
        [string]$Name,

        # Description of the group.
        [string]$Description,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    $Body = @{
        name = $Name
        description = $Description
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
        Invoke-FreshAPIPut -path "requester_groups/$RequesterGroupID" -field "requester_group" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requester_group_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Remove-FreshRequesterGroup {
    <#
    .SYNOPSIS
        This function deletes a requester group
    .DESCRIPTION
        This function deletes a requester group.
    .OUTPUTS
        None.
    .INPUTS
        Object. A representation of the requester group to delete - requires the requester_group_id property.
    .EXAMPLE
        Get-FreshRequesterGroup | Where-Object Name -like "Finance*" | Remove-FreshRequesterGroup
        Deletes all requester groups whose names begin with 'Finance'
    #>
    [CmdletBinding()]
    param(
        # Fresh requester group ID to delete
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_group_id')]
        [int64]$RequesterGroupID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            (Invoke-FreshAPIDelete -path "requester_groups/$RequesterGroupID" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')).group
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

function Add-FreshRequesterGroupMember {
    <#
    .SYNOPSIS
        Adds a Fresh requester to a requester group
    .DESCRIPTION
        Adds a Fresh requester to a requester group
    .NOTES
        Requesters can be added only to manual requester groups. Requesters can be added one at a time.
        Agents cannot be added to requester groups.
    .INPUTS
        Object. Representation of a Fresh requester - must include the requester_id property.
    .OUTPUTS
        None.
    .EXAMPLE
        Get-FreshRequester -email captain.scarlet@spectrum.org | Add-FreshRequesterGroupMember -RequesterGroupID 47
        Adds Captain Scarlet to the requester group with id 47.
    #> 
    [CmdletBinding()]
    param(
        # Name of the group.
        [parameter(Mandatory=$true)]
        [Alias('requester_group_id')]
        [string]$RequesterGroupID,

        # Member to add
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment       
    )

    process {
        try {
            Invoke-FreshAPIPost -path "requester_groups/$RequesterGroupID/members/$RequesterID" -field "requester_group" -body '' -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') #| Select-Object *,@{name='requester_group_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }    
}

function Remove-FreshRequesterGroupMember {
    <#
    .SYNOPSIS
        This function removes a requester from a group
    .DESCRIPTION
        This function deletes a requester group.
    .INPUTS
        Object. A representation of a requester - must include the requester_id
    .OUTPUTS
        None.
    .EXAMPLE
        Get-FreshRequester -email captain.scarlet@spectrum.org | Remove-FreshRequesterGroupMember -RequesterGroupID 47
        Removes Captain Scarlet from the requester group with id 47.
    #>
    [CmdletBinding()]
    param(
        # Fresh requester group ID to delete
        [parameter(Mandatory=$true)]
        [Alias('requester_group_id')]
        [int64]$RequesterGroupID,

        # Fresh Requester ID to remove
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            Invoke-FreshAPIDelete -path "requester_groups/$RequesterGroupID/members/$RequesterID" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')
        } catch {
            $_ | Convert-FreshError
        }        
    }    
}

function Get-FreshRequesterGroupMember {
    <#
    .SYNOPSIS
        Retrieves a list of Fresh requester group members
    .DESCRIPTION
        Retrieves a list of Fresh requester group members
    .INPUTS
        Object. A representation of the requester group -must include the requester_group_id property.
    .OUTPUTS
        Object[]. An array of objects representing the requesters
    .EXAMPLE
        Get-FreshRequesterGroup | Where-Object Name -eq 'Directors' | Get-FreshRequesterGroupMember
        Retrieves the list of members of the 'Directors' requester group.
    #>
    [CmdletBinding()]
    param(
        # ID of requester group
        [parameter(Mandatory=$true,
            ParameterSetName="requester_group",
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requester_group_id')]
        $RequesterGroupID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this ensures the -Verbose setting is passed through to subsequent functions
    }

    process {
        try {
            Invoke-FreshAPIGet -path "requester_groups/$RequesterGroupID/members" -field "requesters" -system $System -verbose:$verbosity | Select-Object *,@{name='requester_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }    
}
