# Module for Fresh ticket related functions.
# Les Newbigging 2022
#
# See the following:
# For Tickets: https://api.freshservice.com/#tickets
# For ticket conversations: https://api.freshservice.com/#conversations

# Create Variables used for tickets

$FreshPriorities = @{}
$FreshStatuses = @{}
$FreshSources = @{}
$FreshUrgencies = @{}
$FreshImpacts = @{}

$FreshSRStages = @{
    'Requested' = 1
    'Delivered' = 2
    'Cancelled' = 3
    'Fulfilled' = 4
    'Partially Fulfilled' = 5
    1 = 'Requested'
    2 = 'Delivered'
    3 = 'Cancelled'
    4 = 'Fulfilled'
    5 = 'Partially Fulfilled'
}

$MandatoryFields = @{}

# INTERNAL SUPPORT FUNCTIONS

function Initialize-FreshVariables {
    <#
    .SYNOPSIS
        Populates the Fresh Hashtables $FreshPriorities, $FreshStatuses, $FreshSources, $FreshUrgencies & $FreshImpacts
    .NOTES
        This will run the first time any of the calls are made, then the hashtables will remain in the environment for later use, reducing the number of API calls made.
    #>
    
    $TicketFields = Get-FreshTicketField 
    $TicketFields | Where-Object name -eq 'priority' | Select-Object -expand choices | foreach-object {$FreshPriorities[[int]$_.id] = $_.Value}
    $TicketFields | Where-Object name -eq 'status' | Select-Object -expand choices | foreach-object {$FreshStatuses[[int]$_.id] = $_.Value}
    $TicketFields | Where-Object name -eq 'urgency' | Select-Object -expand choices | foreach-object {$FreshUrgencies[[int]$_.id] = $_.Value}
    $TicketFields | Where-Object name -eq 'impact' | Select-Object -expand choices | foreach-object {$FreshImpacts[[int]$_.id] = $_.Value}
    $TicketFields | Where-Object name -eq 'source' | Select-Object -expand choices | foreach-object {$FreshSources[[int]$_.id] = $_.Value}
}

function Initialize-FreshTicketMandatoryFields {
    <#
    .SYNOPSIS
        Initialises the MandatoryFields hashtable
    .DESCRIPTION
        Initialises the MandatoryFields hashtable, creating the keys ready for use.
    #>
    param(
        # Which Fresh system/environment to query
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment          
    )
    $AllFields = Get-FreshTicketField -system $System
    $mFields = $AllFields | Where-Object {$_.required_for_agents -or $_.required_for_customers} | Select-Object -ExpandProperty name
    foreach ($mField in $mFields)
    {
        $MandatoryFields.Add($mField,$null)
    }
}

function Get-FreshSRStage {
    <#
    .SYNOPSIS
        Converts the Service request stage number to more readable text
    #>
    param(
        [int]$stage
    )
    return $FreshSRStages[$stage]
}

function Get-FreshTicketPriority {
    <#
    .SYNOPSIS
        Converts the ticket priority number to more readable text
    #>
    param(
        [int]$priority
    )

    if ($FreshPriorities.Keys.count -eq 0)
    {
        Initialize-FreshVariables
    }
    $FreshPriorities[$priority]
}

function Get-FreshTicketStatus {
    <#
    .SYNOPSIS
        Converts the ticket status number to more readable text
    #>
    param(
        [int]$status
    )
    if ($FreshStatuses.Keys.Count -eq 0)
    {
        Initialize-FreshVariables
    }
    $FreshStatuses[$status]
}

function Get-FreshTicketUrgency {
    <#
    .SYNOPSIS
        Converts the ticket urgency number to more readable text
    #>
    param(
        [int]$urgency
    )
    if ($FreshUrgencies.Keys.Count -eq 0)
    {
        Initialize-FreshVariables 
    }
    $FreshUrgencies[$urgency]    
}

function Get-FreshTicketImpact {
    <#
    .SYNOPSIS
        Converts the ticket impact number to more readable text
    #>
    param(
        [int]$impact
    )
    if ($FreshImpacts.Keys.count -eq 0)
    {
        Initialize-FreshVariables
    }
    $FreshImpacts[$impact]
}

function Get-FreshTicketSource {
    <#
    .SYNOPSIS
        Converts the ticket source number to more readable text
    #>
    param(
        [int]$source
    )

    if ($FreshSources.Keys.count -eq 0)
    {
        Initialize-FreshVariables
    }
    $FreshSources[$source]
}

# TICKET RELATED FUNCTIONS

function New-FreshTicket {
    <#
    .SYNOPSIS
        Creates a new Fresh Ticket
    .DESCRIPTION
        Creates a new Fresh Ticket.
    .NOTES
        Information or caveats about the function e.g. 'This function is not supported in Linux'
    .EXAMPLE
        New-FreshTicket -Description "Details about the issue..." -Subject "Support Needed..." -Email "tom@outerspace.com" -Priority 1 -status 2 -CCEmails "ram@freshservice.com","diana@freshservice.com"
        Creates a ticket
    .EXAMPLE
        New-FreshTicket -Description "Details about the issue..." -Subject "Support Needed..." -Email "tom@outerspace.com" -Priority 1 -Status  2 -CCEmails "ram@freshservice.com","diana@freshservice.com" -CustomFields @{ custom_text = "This is a custom text box" } 
        Creates a ticket with custom fields
    .EXAMPLE
        New-FreshTicket -Attachments 'C:/Users/user/Desktop/api_attach.png' -Subject 'Support Needed...' -Description 'Details about the issue...' -Email 'tom@outerspace.com' -Priority 1 -status 2
        Creates a ticket with an attachment
    .EXAMPLE
        New-FreshTicket -Description "Create ticket with assets linked"-Status 2 -Email"sample@freshservice.com" -Priority 1 -Subject "Create ticket with assets linked" -Assets @{ "display_id" = 8 },@{ "display_id" = 9 }
        Creates a ticket with assets
    .EXAMPLE
        New-FreshTicket -Description "Details about the issue..." -Subject "Support Needed..." -Email "tom@outerspace.com" -Priority 1 -Status 2 -Problem @{"display_id" = 2} -ChangeInitiatingTicket @{"display_id" = 4} -ChangeInitiatedByTicket @{"display_id" = 5} 
        Creates a ticket with associations.
    .INPUTS
        None. Does not accept pipeline input.
    .OUTPUTS
        Object
    .NOTES
        Although 'type' is included as a parameter in the API, it has not been included as the only valid value currently is 'Incident', making it a redundant setting.
    #>
    [CmdletBinding(DefaultParameterSetName='phone')]
    param(
        # Name of the requester
        [parameter(Mandatory=$true,
            ParameterSetName='phone')]
        [parameter(Mandatory=$false,
            ParameterSetName='email')]
        [parameter(Mandatory=$false,
            ParameterSetName='requester_id')]    
        [string]$Name,

        # User ID of the requester. For existing contacts, the requester_id can be passed instead of the requester's email.
        [parameter(Mandatory=$true,
            ParameterSetName='requester_id')]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Email address of the requester. If no contact exists with this email address in Freshservice, it will be added as a new contact.
        [parameter(Mandatory=$true,
            ParameterSetName='email')]
        [parameter(Mandatory=$false,
            ParameterSetName='requester_id')]            
        [string]$Email,

        # Phone number of the requester. If no contact exists with this phone number in Freshservice, it will be added as a new contact. If the phone number is set and the email address is not, then the name attribute is mandatory.
        [parameter(Mandatory=$true,
            ParameterSetName='phone')]
        [parameter(Mandatory=$false,
            ParameterSetName='email')]
        [parameter(Mandatory=$false,
            ParameterSetName='requester_id')]                        
        [string]$Phone,

        # Subject of the ticket.
        [string]$Subject,

        # Status of the ticket. 2 = Open (default), 3 = Pending, 4 = Resolved, 5 = Closed
        [int]$Status = 2,

        # Priority of the ticket. 1 = Low (default), 2 = Medium, 3 = High, 4 = Urgent
        [int]$Priority = 1,

        # HTML content of the ticket.
        [string]$Description,

        # ID of the agent to whom the ticket has been assigned
        [Alias('responder_id')]
        [int64]$ResponderID, 

        # Ticket attachments. The total size of these attachments cannot exceed 15MB.
        [string[]]$Attachments,

        # Email address added in the 'cc' field of the incoming ticket email.
        [Alias('cc_emails')]
        [string[]]$CCEmails,
        
        # Key value pairs (hash table) containing the names and values of custom fields. Read more at https://support.freshservice.com/en/support/solutions/articles/154126-customizing-ticket-fields
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Timestamp that denotes when the ticket is due to be resolved.
        [Alias('due_by')]
        [datetime]$DueBy,

        # ID of email config which is used for this ticket. (i.e., support@yourcompany.com/sales@yourcompany.com)
        [Alias('email_config_id')]
        $EmailConfigID,

        # Timestamp that denotes when the first response is due
        [Alias('fr_due_by','FRDueBy')]
        [datetime]$FirstResponseDueBy,

        # ID of the group to which the ticket has been assigned. The default value is the ID of the group that is associated with the given email_config_id
        [Alias('group_id')]
        [int64]$GroupID,

        # The channel through which the ticket was created. The default value is 2 (Portal).
        [int]$Source = 2,        
        
        # Tags to be associated with the ticket
        [string[]]$Tags,

        # Department ID of the requester.
        [Alias('department_id','DeptID')]
        [int64]$DepartmentID,

        # Ticket category
        [string]$Category,

        # Ticket sub category
        [Alias('sub_category')]
        [string]$SubCategory,

        # Ticket item category
        [Alias('item_category')]
        [string]$ItemCategory,

        # Assets that have to be associated with the ticket
        [hashtable]$Assets,

        # Ticket urgency. 1 = Low (default), 2 = Medium, 3 = High
        [int]$Urgency = 1,

        # Ticket impact. 1 = Low (default), 2 = Medium, 3 = High
        [int]$Impact = 1,

        # Problem that need to be associated with ticket (problem display id)
        [hashtable]$Problem,

        # Change causing the ticket that needs to be associated with ticket (change display id)
        [Alias('change_initiating_ticket')]
        [hashtable]$ChangeInitiatingTicket,

        # Change needed for the ticket to be fixed that needs to be associated with ticket (change display id)
        [Alias('change_initiated_by_ticket')]
        [hashtable]$ChangeInitiatedByTicket,

        # Which Fresh system/environment to query
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue'
    }

    process {

        $parameterHash = @{}

        if ($Name)
        {
            $parameterHash.Add("name",$Name)
        }
        
        if ($RequesterID)
        {
            $parameterHash.Add("requester_id",$RequesterID)
        }

        if ($Email)
        {
            $parameterHash.Add("email",$Email)
        }

        if ($Phone)
        {
            $parameterHash.Add("phone",$phone)
        }

        if ($Subject)
        {
            $parameterHash.Add("subject",$Subject)
        }

        if ($Status)
        {
            $parameterHash.Add("status",$Status)
        }

        if ($Priority)
        {
            $parameterHash.Add("priority",$Priority)
        }
        
        if ($Description)
        {
            $parameterHash.Add("description",$Description)
        }

        if ($ResponderID)
        {
            $parameterHash.Add("responder_id",$ResponderID)
        }

        if ($Attachments)
        {
            $parameterHash.Add("attachments[]",$Attachments)
        }

        if($CCEmails)
        {
            $parameterHash.Add("cc_emails",$CCEmails)
        }

        if ($CustomFields)
        {
            $parameterHash.Add("custom_fields",$CustomFields)
        }

        if ($DueBy)
        {
            $parameterHash.Add("due_by",(ConvertTo-ISO8601Date -DateTime $DueBy))
        }

        if ($EmailConfigID)
        {
            $parameterHash.Add("email_config_id",$EmailConfigID)
        }

        if ($FirstResponseDueBy)
        {
            $parameterHash.Add("fr_due_by",(ConvertTo-ISO8601Date -DateTime $FirstResponseDueBy))
        }

        if ($GroupID)
        {
            $parameterHash.Add("group_id",$GroupID)
        }

        if ($Source)
        {
            $parameterHash.Add("source",$Source)
        }

        if ($Tags)
        {
            $parameterHash.Add("tags",$Tags)
        }

        if ($DepartmentID)
        {
            $parameterHash.Add("department_id",$DepartmentID)
        }

        if ($Category)
        {
            $parameterHash.Add("category",$Category)
        }

        if ($SubCategory)
        {
            $parameterHash.Add("sub_category",$SubCategory)
        }

        if ($ItemCategory)
        {
            $parameterHash.Add("item_category",$ItemCategory)
        }

        if ($Assets)
        {
            $parameterHash.Add("assets",$assets)
        }

        if ($Urgency)
        {
            $parameterHash.Add("urgency",$Urgency)
        }

        if ($Impact)
        {
            $parameterHash.Add("impact",$Impact)
        }

        if ($Problem)
        {
            $parameterHash.Add("problem",$Problem)
        }

        if ($ChangeInitiatingTicket)
        {
            $parameterHash.Add("change_initiating_ticket",$ChangeInitiatingTicket)
        }

        if ($ChangeInitiatedByTicket)
        {
            $parameterHash.Add("change_initiated_by_ticket",$ChangeInitiatedByTicket)
        }

        # Check that all required fields are completed
        if ($MandatoryFields.count -eq 0)
        {
            Initialize-FreshTicketMandatoryFields
            Write-Verbose "Configured Mandatory fields: $($MandatoryFields.Keys -join ',')"
        }

        $MissingFields = $MandatoryFields.Keys | Where-Object {$_ -notin $parameterHash.Keys -and $_ -ne 'requester'}
        if (($MissingFields | measure-object).count -ne 0)
        {
            throw "The following fields are reported as MANDATORY: $($MissingFields -join ', '). Please ensure all required parameters are provided."
        }

        # If this has attachments, it needs to use multipart/form
        if ($parameterHash.Keys -contains "attachments[]")
        {
            # Includes attachments
            $boundary = [System.Guid]::NewGuid().ToString()
            $ContentType = "multipart/form-data; boundary=`"$boundary`"" 
            $PostBody = ConvertTo-MultiPartFormData -parameters $parameterHash -boundary $boundary
        } else {
            # No attachments - use application/json
            $ContentType = 'application/json'
            $PostBody = $parameterHash | Convertto-Json -depth 5
        }

        try {
            Invoke-FreshAPIPost -path "tickets" -field 'ticket' -body $PostBody -ContentType $ContentType -System $System -Verbose:$verbosity | Select-Object *,@{name='status_text';exp={Get-FreshTicketStatus $_.status}},@{name='priority_text';exp={Get-FreshTicketPriority $_.priority}},@{name='source_text';exp={Get-FreshTicketSource $_.source}},@{name='ticket_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }

}

function Get-FreshTicket {
    <#
    .SYNOPSIS
        Retrieves ticket(s) from Fresh
    .DESCRIPTION
        Retrieves tickets from Fresh, by ID, by filter conditions, or all.
    .EXAMPLE
        Get-FreshTicket
        Returns all fresh tickets that are neither spam nor deleted
    .EXAMPLE
        Get-FreshTicket -TicketID 42
        Returns the details for ticket #42
    .EXAMPLE
        Get-FreshTicket -PreDefinedFilter deleted
        Returns all deleted tickets
    .EXAMPLE
        Get-FreshTicket -Priority Urgent,High -Status Open  
        Returns all Open tickets with Urgent or High priority
    .EXAMPLE
        Get-FreshTicket -Tag 'Awaiting 3rd Party' -IncludeTags 
        Retrieves all tickets with the 'Awaiting 3rd Party' tag, and returns the 'tags' fields along with the tickets
    .INPUTS
        Object. Must include ticket_id as a property
    .OUTPUTS
        Object[]
    .NOTES
        Includes textualised properties for impact, priority, source, status & urgency for readability.
    #>
    [cmdletBinding(DefaultParameterSetName='tickets')]
    param(
        # Finds a single Fresh Ticket by (numeric) ID
        [parameter(Mandatory=$true,
            ParameterSetName='ticket',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Search by ID of the agent to whom the ticket has been assigned
        [parameter(ParameterSetName='filter')]
        [Alias('agent_id')]
        [int64]$AgentID,

        # Search by ID of the group to which the ticket has been assigned
        [parameter(ParameterSetName='filter')]
        [Alias('group_id')]
        [int64]$GroupID,

        # Search by priority of the ticket [Low|Medium|High|Urgent]
        [parameter(ParameterSetName='filter')]
        [ValidateSet('Low','Medium','High','Urgent')]
        [string[]]$Priority,

        # Search by status of the ticket [Open|Pending|Resolved|Closed]
        [parameter(ParameterSetName='filter')]
        [ValidateSet('Open','Pending','Resolved','Closed')]
        [string[]]$Status,

        # Search by ticket impact rating (numeric)
        [parameter(ParameterSetName='filter')]
        [int]$Impact,

        # Search by ticket urgency (numberic)
        [parameter(ParameterSetName='filter')]
        [int]$Urgency,

        # Search for tickets with a particular tag 
        [parameter(ParameterSetName='filter')]
        [string]$Tag,

        # Pre-defined filters to use [new_and_my_open|watching|spam|deleted]
        [parameter(ParameterSetName='tickets')]
        [ValidateSet('new_and_my_open','watching','spam','deleted')]
        [string]$PreDefinedFilter,

        # Search by Fresh ID of the requester rasing the ticket
        [parameter(ParameterSetName='tickets')]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Search by requester's email address (primary and secondary)
        [parameter(ParameterSetName='tickets')]
        [Alias('primary_email')]
        [string]$Email,

        # Search by ticket type [Incident|Service Request]
        [parameter(ParameterSetName='tickets')]
        [ValidateSet('Incident','Service Request')]
        [string]$Type,

        # Include Conversations from the ticket (first 10 only). For more, use Get-FreshTicketConversation with Ticket ID
        [parameter(parametersetname='ticket')]
        [Alias('conversations')]
        [switch]$IncludeConversations,

        # Include expanded requester details from the tickets
        [Alias('requester')]
        [switch]$IncludeRequester,

        # Include requested for details from the tickets
        [Alias('requested_for')]
        [switch]$IncludeRequestedFor,

        # Include Department details in the tickets
        [switch]$IncludeDepartment,

        # Include tags with the tickets
        [switch]$IncludeTags,

        # Include ticket stats with the results
        [Alias('stats')]
        [switch]$IncludeStats,

        # Include associated problem details with the ticket
        [parameter(parametersetname='ticket')]
        [Alias('problem')]
        [switch]$IncludeProblem,

        # Include associated asset details with the ticket
        [parameter(parametersetname='ticket')]
        [Alias('assets','asset')]
        [switch]$IncludeAssets,

        # Include associated change details
        [parameter(parametersetname='ticket')]
        [Alias('change')]
        [switch]$IncludeChange,

        # Include related ticket ids
        [parameter(parametersetname='ticket')]
        [Alias('related_tickets','relatedtickets')]
        [switch]$IncludeRelatedTickets,        

        # Which Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment
    )
    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue' # this enmsure the -Verbose setting is passed through to subsequent functions
    }

    process {
        $path = "tickets"
        $field = $PSCmdlet.ParameterSetName
        write-verbose "ParameterSetName: $field"

        $FilterParameters = @()

        switch($field)
        {
            'ticket' {
                $path += "/$TicketID"
            }

            'tickets' {
                # 100 is the maximum allowed page size, and will minimise the numer of calls made to the API, as the default is 30.
                $FilterParameters += "per_page=100"
            }

            'filter' {
                $field = "tickets"
                $path += "/filter"
                $FilterParameters += "per_page=100"
            }
        }

        $Includes = @()
        if ($IncludeConversations)
        {
            $Includes += "conversations"
        }

        if ($IncludeRequester)
        {
            $Includes += "requester"
        }

        if ($IncludeRequestedFor)
        {
            $Includes += "requested_for"
        }

        if ($IncludeDepartment)
        {
            $Includes += "department"
        }

        if ($IncludeTags)
        {
            $Includes += "tags"
        }

        if ($IncludeStats)
        {
            $Includes += "stats"
        }

        if ($IncludeProblem)
        {
            $Includes += "problem"
        }

        if ($IncludeAssets)
        {
            $Includes += "assets"
        }

        if ($IncludeChange)
        {
            $Includes += "change"
        }

        if ($IncludeRelatedTickets)
        {
            $Includes += "related_tickets"
        }

        if ($Includes.count -ne 0)
        {
            $FilterParameters += "include=" + ($Includes -join ',')
        }

        $Queries = @()

        if ($AgentID)
        {
            $Queries += "agent_id:$AgentID"
        }

        if ($GroupID)
        {
            $Queries += "group_id:$GroupID"
        }

        if ($Priority)
        {
            $PriorityArray = @()
            foreach ($p in $priority)
            {
                $PriorityArray += "priority:$($FreshPriorities[$p])"
            }
            $Queries += "(" + ($PriorityArray -join ' OR ') + ")"
        }

        if ($Status)
        {
            $StatusArray = @()
            foreach ($s in $Status)
            {
                $StatusArray += "status:$($FreshStatuses[$s])"
            }
            $Queries += "(" + ($StatusArray -join ' OR ') + ")"
        }    

        if ($Impact)
        {
            $Queries += "impact:$Impact"
        }
        
        if ($Urgency)
        {
            $Queries += "urgency:$urgency"
        }
        
        if ($Tag)
        {
            $Queries += "tag:'$Tag'"
        }

        if ($Queries.count -ne 0)
        {
            $FilterParameters += 'query="' + ($Queries -join ' AND ') + '"'
        }

        if ($PreDefinedFilter)
        {
            $FilterParameters += "filter=$PreDefinedFilter"
        }

        if ($RequesterID)
        {
            $FilterParameters += "requester_id=$RequesterID"
        }

        if ($Email)
        {
            $FilterParameters += "email=$Email"
        }

        if ($Type)
        {
            $FilterParameters += "type=$Type"
        }

        if ($FilterParameters.count -ne 0)
        {
            $path += "?" + ($FilterParameters -join '&')
        }

        try {
            Invoke-FreshAPIGet -path $path -field $field -system $System -verbose:$verbosity | Select-Object *,@{name='status_text';exp={Get-FreshTicketStatus $_.status}},@{name='priority_text';exp={Get-FreshTicketPriority $_.priority}},@{name='source_text';exp={Get-FreshTicketSource $_.source}},@{name='urgency_text';exp={Get-FreshTicketUrgency $_.urgency}},@{name='impact_text';exp={Get-FreshTicketImpact $_.impact}},@{name='ticket_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Update-FreshTicket {
    <#
    .SYNOPSIS
        Make changes to the parameters of a ticket 
    .DESCRIPTION
        This function lets you make changes to the parameters of a ticket from updating statuses to changing ticket type.
    .EXAMPLE
        Update-FreshTicket -Description "Update ticket with assets" -Status 2 -Email "sample@freshservice.com" -Priority 1 -Subject "Update ticket with assets" -Assets @{ "display_id": 7 },@{ "display_id": 8 }
        Updates a ticket with assets;  Existing assets, if they are different from what are given in request, are destroyed and the current ones are linked to the ticket. So, all the assets that need to stay associated with the Ticket need to be provided in the details.
    .EXAMPLE
        Update-FreshTicket -Attachments 'C:/Users/user/Desktop/api2.png' -Priority 1
        Updates a ticket with an attachment
    .INPUTS
        Object. The ticket to update - requires the ticket_id property.
    .OUTPUTS
        Object. The updated ticket.
    .NOTES
        Although 'type' is included as a parameter in the API, it has not been included in this function, as the only valid value currently is 'Incident', rendering it a redundant setting.
    #>
    [CmdletBinding()]
    param(
        # ID of Fresh ticket to update
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        $TicketID,

        # Name of the requester
        [string]$Name,

        # User ID of the requester. For existing contacts, the requester_id can be passed instead of the requester's email.
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Email address of the requester. If no contact exists with this email address in Freshservice, it will be added as a new contact.
        [string]$Email,

        # Phone number of the requester. If no contact exists with this phone number in Freshservice, it will be added as a new contact. If the phone number is set and the email address is not, then the name attribute is mandatory.
        [string]$Phone,

        # Subject of the ticket.
        [string]$Subject,

        # Type helps categorize the ticket according to the different kinds of issues your support team deals with
        [ValidateSet('Incident','Service Request')]
        [string]$Type, 

        # Status of the ticket. 2 = Open, 3 = Pending, 4 = Resolved, 5 = Closed
        [int]$Status,

        # Priority of the ticket. 1 = Low, 2 = Medium, 3 = High, 4 = Urgent
        [int]$Priority,

        # HTML content of the ticket.
        [string]$Description,

        # ID of the agent to whom the ticket has been assigned
        [Alias('responder_id')]
        [int64]$ResponderID, 

        # Ticket attachments. The total size of these attachments cannot exceed 15MB.
        [string[]]$Attachments,

        # Key value pairs (hash table) containing the names and values of custom fields. Read more at https://support.freshservice.com/en/support/solutions/articles/154126-customizing-ticket-fields
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Timestamp that denotes when the ticket is due to be resolved.
        [Alias('due_by')]
        [datetime]$DueBy,

        # ID of email config which is used for this ticket. (i.e., support@yourcompany.com/sales@yourcompany.com)
        [Alias('email_config_id')]
        $EmailConfigID,

        # Timestamp that denotes when the first response is due
        [Alias('fr_due_by','FRDueBy')]
        [datetime]$FirstResponseDueBy,

        # ID of the group to which the ticket has been assigned. The default value is the ID of the group that is associated with the given email_config_id
        [Alias('group_id')]
        [int64]$GroupID,

        # The channel through which the ticket was created. The default value is 2 (Portal).
        [int]$Source,        
        
        # Tags to be associated with the ticket
        [string[]]$Tags,

        # Assets that have to be associated with the ticket
        [hashtable]$Assets,

        # Ticket urgency. 1 = Low 2 = Medium, 3 = High
        [int]$Urgency,

        # Ticket impact. 1 = Low, 2 = Medium, 3 = High
        [int]$Impact,

        # Ticket category
        [string]$Category,

        # Ticket sub category
        [Alias('sub_category')]
        [string]$SubCategory,

        # Ticket item category
        [Alias('item_category')]
        [string]$ItemCategory,

        # Problem that need to be associated with ticket (problem display id)
        [hashtable]$Problem,

        # Change causing the ticket that needs to be associated with ticket (change display id)
        [Alias('change_initiating_ticket')]
        [hashtable]$ChangeInitiatingTicket,

        # Change needed for the ticket to be fixed that needs to be associated with ticket (change display id)
        [Alias('change_initiated_by_ticket')]
        [hashtable]$ChangeInitiatedByTicket,

        # Which Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue'
    }

    process {

        $parameterHash = @{}

        if ($Name)
        {
            $parameterHash.Add("name",$Name)
        }

        if ($RequesterID)
        {
            $parameterHash.Add("requester_id",$RequesterID)
        }

        if ($Email)
        {
            $parameterHash.Add("email",$Email)
        }

        if ($Phone)
        {
            $parameterHash.Add("phone",$phone)
        }

        if ($Subject)
        {
            $parameterHash.Add("subject",$Subject)
        }

        if ($Type)
        {
            $parameterHash.Add("type",$Type)
        }

        if ($Status)
        {
            $parameterHash.Add("status",$Status)
        }

        if ($Priority)
        {
            $parameterHash.Add("priority",$Priority)
        }
        
        if ($Description)
        {
            $parameterHash.Add("description",$Description)
        }

        if ($ResponderID)
        {
            $parameterHash.Add("responder_id",$ResponderID)
        }

        if ($Attachments)
        {
            $parameterHash.Add("attachments[]",$Attachments)
        }

        if ($CustomFields)
        {
            $parameterHash.Add("custom_fields",$CustomFields)
        }

        if ($DueBy)
        {
            $parameterHash.Add("due_by",(ConvertTo-ISO8601Date -DateTime $DueBy))
        }

        if ($EmailConfigID)
        {
            $parameterHash.Add("email_config_id",$EmailConfigID)
        }

        if ($FirstResponseDueBy)
        {
            $parameterHash.Add("fr_due_by",(ConvertTo-ISO8601Date -DateTime $FirstResponseDueBy))
        }

        if ($GroupID)
        {
            $parameterHash.Add("group_id",$GroupID)
        }

        if ($Source)
        {
            $parameterHash.Add("source",$Source)
        }

        if ($Tags)
        {
            $parameterHash.Add("tags",$Tags)
        }

        if ($Category)
        {
            $parameterHash.Add("category",$Category)
        }

        if ($SubCategory)
        {
            $parameterHash.Add("sub_category",$SubCategory)
        }

        if ($ItemCategory)
        {
            $parameterHash.Add("item_category",$ItemCategory)
        }

        if ($Assets)
        {
            $parameterHash.Add("assets",$assets)
        }

        if ($Urgency)
        {
            $parameterHash.Add("urgency",$Urgency)
        }

        if ($Impact)
        {
            $parameterHash.Add("impact",$Impact)
        }

        if ($Problem)
        {
            $parameterHash.Add("problem",$Problem)
        }

        if ($ChangeInitiatingTicket)
        {
            $parameterHash.Add("change_initiating_ticket",$ChangeInitiatingTicket)
        }

        if ($ChangeInitiatedByTicket)
        {
            $parameterHash.Add("change_initiated_by_ticket",$ChangeInitiatedByTicket)
        }

        # If this has attachments, it needs to use multipart/form
        if ($parameterHash.Keys -contains "attachments[]")
        {
            # Includes attachments
            $boundary = [System.Guid]::NewGuid().ToString()
            $ContentType = "multipart/form-data; boundary=`"$boundary`"" 
            $PostBody = ConvertTo-MultiPartFormData -parameters $parameterHash -boundary $boundary
        } else {
            # No attachments - use application/json
            $ContentType = 'application/json'
            $PostBody = $parameterHash | Convertto-Json -depth 5
        }

        try {
            Invoke-FreshAPIPut -path "tickets/$TicketID" -field 'ticket' -body $PostBody -ContentType $ContentType -System $System -Verbose:$verbosity | Select-Object *,@{name='status_text';exp={Get-FreshTicketStatus $_.status}},@{name='priority_text';exp={Get-FreshTicketPriority $_.priority}},@{name='source_text';exp={Get-FreshTicketSource $_.source}},@{name='ticket_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }

}

function Remove-FreshTicket {
    <#
    .SYNOPSIS
        Deletes a ticket
    .DESCRIPTION
        Deletes a ticket. Deleted tickets can be restored.
    .NOTES
        Trying to delete an already deleted ticket results in a 405 - Method not allowed (DELETE)
    .EXAMPLE
        Remove-FreshTicket -TicketID 123
        Deletes ticket 123
    .EXAMPLE
        Get-FreshTicket -Status Closed | Remove-FreshTicket
        Deletes all closed tickets.
    .INPUTS
        Object. Can be a list of tickets genereated from Get-FreshTicket
    .OUTPUTS
        None. 
    #>
    [CmdletBinding()]
    param(
        # FreshTicket (numeric) ID
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName)]        
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment          
    )

    process {
        try {
            Invoke-FreshAPIDelete -path "tickets/$TicketID" -system $System -verbose:($VerbosePreference -ne 'SilentlyContinue')
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Remove-FreshTicketAttachment {
    <#
    .SYNOPSIS
        This deletes an attachment from a ticket.
    .DESCRIPTION
        This deletes an attachment from a ticket.
    .EXAMPLE
        Remove-FreshTicketAttachment -TicketID 123 -AttachmentID 2
        Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
    #>
    [cmdletBinding()]
    param(
        # FreshTicket (numeric) ID
        [parameter(Mandatory=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # ID of attachment to remove
        [parameter(Mandatory=$true)]
        $AttachmentID,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    try {
        Invoke-FreshAPIDelete -path "tickets/$TicketID/attachments/$AttachmentID" -system $System -verbose:($VerbosePreference -ne 'SilentlyContinue')
    } catch {
        $_ | Convert-FreshError
    }
}

function Restore-FreshTicket {
    <#
    .SYNOPSIS
        Restore a deleted ticket
    .DESCRIPTION
        Quoting the API documentation:  If you deleted some tickets and regret doing so now, this API will help you restore them.
    .NOTES
        If using in a script, it may be useful to verify the deleted ticket ID first using 'Get-FreshTicket -PreDefinedFilter deleted', 
        or by directly querying the ticket by ID and checking the "deleted" property.
    .EXAMPLE
        Restore-FreshTicket -TicketID 5
        Restores a deleted ticket with ID of 5.
    .INPUTS
        Object. Must include the ticket_id property.
    .OUTPUTS
        None.
    #>
    [CmdletBinding()]
    param(
        # FreshTicket (numeric) ID
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true,
            position=0)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    process {
        try {
            Invoke-FreshAPIPut -path "tickets/$TicketID/restore" -verbose:($VerbosePreference -ne 'SilentlyContinue')
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function New-FreshChildTicket {
    <#
    .SYNOPSIS
        Creates a new Fresh Child Ticket
    .DESCRIPTION
        Creates a new Fresh Child Ticket.
    .EXAMPLE
        New-FreshChildTicket -ParentTicketID 123 -Description "Details about the issue..." -Subject "Support Needed..." -Email "tom@outerspace.com" -Priority 1 -status 2 -CCEmails "ram@freshservice.com","diana@freshservice.com"
        Creates a child ticket from ticket 123
    .INPUTS
        Object. The parent ticket
    .OUTPUTS
        Object. The child ticket.
    .NOTES
        This will only work with INCIDENT tickets, creating an incident ticket.
    #>
    [CmdletBinding(DefaultParameterSetName='phone')]
    param(
        # Parent Ticket ID
        [parameter(Mandatory=$true)]
        [Alias('ticket_id','TicketID')]
        $ParentTicketID,

        # Name of the requester
        [parameter(Mandatory=$true,
            ParameterSetName='phone')]
        [parameter(Mandatory=$false,
            ParameterSetName='email')]
        [parameter(Mandatory=$false,
            ParameterSetName='requester_id')]    
        [string]$Name,

        # User ID of the requester. For existing contacts, the requester_id can be passed instead of the requester's email.
        [parameter(Mandatory=$true,
            ParameterSetName='requester_id')]
        [Alias('requester_id')]
        [int64]$RequesterID,

        # Email address of the requester. If no contact exists with this email address in Freshservice, it will be added as a new contact.
        [parameter(Mandatory=$true,
            ParameterSetName='email')]
        [parameter(Mandatory=$false,
            ParameterSetName='requester_id')]            
        [string]$Email,

        # Phone number of the requester. If no contact exists with this phone number in Freshservice, it will be added as a new contact. If the phone number is set and the email address is not, then the name attribute is mandatory.
        [parameter(Mandatory=$true,
            ParameterSetName='phone')]
        [parameter(Mandatory=$false,
            ParameterSetName='email')]
        [parameter(Mandatory=$false,
            ParameterSetName='requester_id')]                        
        [string]$Phone,

        # Subject of the ticket.
        [string]$Subject,

        # Status of the ticket. 2 = Open (default), 3 = Pending, 4 = Resolved, 5 = Closed
        [int]$Status = 2,

        # Priority of the ticket. 1 = Low (default), 2 = Medium, 3 = High, 4 = Urgent
        [int]$Priority = 1,

        # HTML content of the ticket.
        [string]$Description,

        # ID of the agent to whom the ticket has been assigned
        [Alias('responder_id')]
        [int64]$ResponderID, 

        # Ticket attachments. The total size of these attachments cannot exceed 15MB.
        [string[]]$Attachments,

        # Email address added in the 'cc' field of the incoming ticket email.
        [Alias('cc_emails')]
        [string[]]$CCEmails,
        
        # Key value pairs (hash table) containing the names and values of custom fields. Read more at https://support.freshservice.com/en/support/solutions/articles/154126-customizing-ticket-fields
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Timestamp that denotes when the ticket is due to be resolved.
        [Alias('due_by')]
        [datetime]$DueBy,

        # ID of email config which is used for this ticket. (i.e., support@yourcompany.com/sales@yourcompany.com)
        [Alias('email_config_id')]
        $EmailConfigID,

        # Timestamp that denotes when the first response is due
        [Alias('fr_due_by','FRDueBy')]
        [datetime]$FirstResponseDueBy,

        # ID of the group to which the ticket has been assigned. The default value is the ID of the group that is associated with the given email_config_id
        [Alias('group_id')]
        [int64]$GroupID,

        # The channel through which the ticket was created. The default value is 2 (Portal).
        [int]$Source = 2,        
        
        # Tags to be associated with the ticket
        [string[]]$Tags,

        # Department ID of the requester.
        [Alias('department_id','DeptID')]
        [int64]$DepartmentID,

        # Ticket category
        [string]$Category,

        # Ticket sub category
        [Alias('sub_category')]
        [string]$SubCategory,

        # Ticket item category
        [Alias('item_category')]
        [string]$ItemCategory,

        # Assets that have to be associated with the ticket
        [hashtable]$Assets,

        # Ticket urgency. 1 = Low (default), 2 = Medium, 3 = High
        [int]$Urgency = 1,

        # Ticket impact. 1 = Low (default), 2 = Medium, 3 = High
        [int]$Impact = 1,

        # Problem that need to be associated with ticket (problem display id)
        [hashtable]$Problem,

        # Change causing the ticket that needs to be associated with ticket (change display id)
        [Alias('change_initiating_ticket')]
        [hashtable]$ChangeInitiatingTicket,

        # Change needed for the ticket to be fixed that needs to be associated with ticket (change display id)
        [Alias('change_initiated_by_ticket')]
        [hashtable]$ChangeInitiatedByTicket,

        # Which Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )

    begin {
        $verbosity = $VerbosePreference -ne 'SilentlyContinue'
    }

    process {

        $parameterHash = @{}

        if ($Name)
        {
            $parameterHash.Add("name",$Name)
        }
        
        if ($RequesterID)
        {
            $parameterHash.Add("requester_id",$RequesterID)
        }

        if ($Email)
        {
            $parameterHash.Add("email",$Email)
        }

        if ($Phone)
        {
            $parameterHash.Add("phone",$phone)
        }

        if ($Subject)
        {
            $parameterHash.Add("subject",$Subject)
        }

        if ($Status)
        {
            $parameterHash.Add("status",$Status)
        }

        if ($Priority)
        {
            $parameterHash.Add("priority",$Priority)
        }
        
        if ($Description)
        {
            $parameterHash.Add("description",$Description)
        }

        if ($ResponderID)
        {
            $parameterHash.Add("responder_id",$ResponderID)
        }

        if ($Attachments)
        {
            $parameterHash.Add("attachments[]",$Attachments)
        }

        if($CCEmails)
        {
            $parameterHash.Add("cc_emails",$CCEmails)
        }

        if ($CustomFields)
        {
            $parameterHash.Add("custom_fields",$CustomFields)
        }

        if ($DueBy)
        {
            $parameterHash.Add("due_by",(ConvertTo-ISO8601Date -DateTime $DueBy))
        }

        if ($EmailConfigID)
        {
            $parameterHash.Add("email_config_id",$EmailConfigID)
        }

        if ($FirstResponseDueBy)
        {
            $parameterHash.Add("fr_due_by",(ConvertTo-ISO8601Date -DateTime $FirstResponseDueBy))
        }

        if ($GroupID)
        {
            $parameterHash.Add("group_id",$GroupID)
        }

        if ($Source)
        {
            $parameterHash.Add("source",$Source)
        }

        if ($Tags)
        {
            $parameterHash.Add("tags",$Tags)
        }

        if ($DepartmentID)
        {
            $parameterHash.Add("department_id",$DepartmentID)
        }

        if ($Category)
        {
            $parameterHash.Add("category",$Category)
        }

        if ($SubCategory)
        {
            $parameterHash.Add("sub_category",$SubCategory)
        }

        if ($ItemCategory)
        {
            $parameterHash.Add("item_category",$ItemCategory)
        }

        if ($Assets)
        {
            $parameterHash.Add("assets",$assets)
        }

        if ($Urgency)
        {
            $parameterHash.Add("urgency",$Urgency)
        }

        if ($Impact)
        {
            $parameterHash.Add("impact",$Impact)
        }

        if ($Problem)
        {
            $parameterHash.Add("problem",$Problem)
        }

        if ($ChangeInitiatingTicket)
        {
            $parameterHash.Add("change_initiating_ticket",$ChangeInitiatingTicket)
        }

        if ($ChangeInitiatedByTicket)
        {
            $parameterHash.Add("change_initiated_by_ticket",$ChangeInitiatedByTicket)
        }

        # Check that all required fields are completed
        if ($MandatoryFields.count -eq 0)
        {
            Initialize-FreshTicketMandatoryFields
            Write-Verbose "Configured Mandatory fields: $($MandatoryFields.Keys -join ',')"
        }

        $MissingFields = $MandatoryFields.Keys | Where-Object {$_ -notin $parameterHash.Keys -and $_ -ne 'requester'}
        if (($MissingFields | measure-object).count -ne 0)
        {
            throw "The following fields are reported as MANDATORY: $($MissingFields -join ', '). Please ensure all required parameters are provided."
        }

        # If this has attachments, it needs to use multipart/form
        if ($parameterHash.Keys -contains "attachments[]")
        {
            # Includes attachments
            $boundary = [System.Guid]::NewGuid().ToString()
            $ContentType = "multipart/form-data; boundary=`"$boundary`"" 
            $PostBody = ConvertTo-MultiPartFormData -parameters $parameterHash -boundary $boundary
        } else {
            # No attachments - use application/json
            $ContentType = 'application/json'
            $PostBody = $parameterHash | Convertto-Json -depth 5
        }

        try {
            Invoke-FreshAPIPost -path "tickets/$ParentTicketID/create_child_ticket" -field 'ticket' -body $PostBody -ContentType $ContentType -System $System -Verbose:$verbosity | Select-Object *,@{name='status_text';exp={Get-FreshTicketStatus $_.status}},@{name='priority_text';exp={Get-FreshTicketPriority $_.priority}},@{name='source_text';exp={Get-FreshTicketSource $_.source}},@{name='ticket_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }

}

function Get-FreshTicketField {
    <#
    .SYNOPSIS
        Retrieve all the Fields that constitute the Ticket Object
    .DESCRIPTION
        Retrieve all the Fields that constitute the Ticket Object, with their properties.
    .EXAMPLE
        Get-FreshTicketField | Select-Object id,name,label,description
        Lists the ticket fields' id, name, label & descriptoin properties.
    .INPUTS
        None. Does not accept pipeline input.
    .OUTPUTS
        Object[]. An array of objects representing the fields.
    #>
    [CmdletBinding()]
    param(
        # Fresh system/environment to query
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment        
    )
    Invoke-FreshAPIGet -path "ticket_form_fields" -field "ticket_fields" -system $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='system';exp={$System}}
}

function Get-FreshTicketActivity {
    <#
    .SYNOPSIS
        Lists the activities performed on a ticket.
    .DESCRIPTION
       This helps to track exactly how much time an agent has spent on each ticket, start/stop timers and perform a lot of other time tracking and monitoring tasks to ensure that the support team is always performing at its peak efficiency.
    .EXAMPLE
        Get-FreshTicketActivity -TicketID 59 
        Lists the actions on ticket id 59
    .INPUTS
        Object. A representation of a Fresh ticket - must include ticket_id as a property.
    .OUTPUTS
        Object[]
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    process {
        try {
            Invoke-FreshAPIGet -path "tickets/$TicketID/activities" -field "activities" -system $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='ticket_id';exp={$TicketID}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

# TIME ENTRY RELATED FUNCTIONS

function New-FreshTicketTimeEntry {
    <#
    .SYNOPSIS
        This function creates a Time Entry.
    .DESCRIPTION
        This function creates a Time Entry.
    .EXAMPLE
        New-FreshTicketTimeEntry -TicketID 3001 -agentID 7007 -TimeSpent "01:30"
        Adds a time entry to say agent 7007 spent 1 1/2 hours on the ticket.
    .EXAMPLE
        New-FreshTicketTimeEntry -TicketID 3002 -agentID 9090 -TimerRunning $true
        Adds a timer entry and sets the clock running.
    .INPUTS
        Object. A Fresh ticket with a ticket_id property.
    .OUTPUTS
        Object. A representation of the time entry.
    #>
    [CmdletBinding()]
    param (
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Set to true if timer is currently running. Default value is false. At a time, only one timer can be running for an agent across the account
        [Alias('timer_running')]
        [AllowNull()]
        [boolean]$TimerRunning,

        # Set as true if the time-entry is billable. Default value is true
        [AllowNull()]
        [boolean]$Billable,

        # The total amount of time spent by the timer in hh:mm format. This field cannot be set if timer_running is true. Mandatory if timer_running is false
        [ValidatePattern("^\d{2}:\d{2}$")]
        [Alias('time_spent')]
        [string]$TimeSpent,

        # Time at which the timer is executed. Default value (unless given in request) is the time at which timer is added. Should be less than or equal to current date_time
        [Alias('executed_at')]
        [datetime]$ExecutedAt,

        # Id of the task assigned to the time-entry. Task should be valid on the given ticket and assigned to agent_id
        [Alias('task_id')]
        [int64]$TaskID,

        # Description of the time-entry
        [string]$Note,

        # The user/agent to whom this time-entry is assigned
        [parameter(Mandatory=$true)]
        [Alias('agent_id')]
        [int64]$AgentID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment         
    )
    
    process {
        $TimeEntry = @{
            timer_running = $TimerRunning
            billable = $Billable
            time_spent = $TimeSpent
            executed_at = Convert-FreshSafeDate $ExecutedAt
            task_id	 = $TaskID
            note = $Note
            agent_id = $AgentID
        }

        # find empty keys
        $EmptyKeys = @()
        foreach ($key in $TimeEntry.keys)
        {
            # A boolean $false will test positive as '' and 0, therefore the -isnot [boolean] test was added
            if ($null -eq $TimeEntry[$key] -or ($TimeEntry[$key] -isnot [boolean] -and ($TimeEntry[$key] -eq '' -or $TimeEntry[$key] -eq 0)))
            {
                $EmptyKeys += $key
            }
        }

        # remove empty keys
        foreach ($key in $EmptyKeys)
        {
            $TimeEntry.Remove($key)
        }

        $jsonBody = @{time_entry=$TimeEntry} | ConvertTo-Json

        try {
            Invoke-FreshAPIPost -path "tickets/$TicketID/time_entries" -field "time_entry" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='ticket_id';exp={$TicketID}},@{name='time_entry_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Get-FreshTicketTimeEntry {
    <#
    .SYNOPSIS
        View all or individual time entries of a particular ticket
    .DESCRIPTION
        View all or individual time entries of a particular ticket
    .INPUTS
        Object. A representation of either a time entry or a Fresh ticket.
    .OUTPUTS
        Object[]. An array of time entry objects.
    #>
    [CmdletBinding(DefaultParameterSetName='time_entries')]
    param(
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh Time Entry ID
        [parameter(Mandatory=$true,
            ParameterSetName='time_entry',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('time_entry_id')]
        $TimeEntryID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )

    begin {
        $field = $PSCmdlet.ParameterSetName
    }

    process {
        $path = "tickets/$TicketID/time_entries"
        if ($TimeEntryID)
        {
            $path += "/$TimeEntryID"
        }
        try {
            Invoke-FreshAPIGet -path $path -field $field -System $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='ticket_id';exp={$TicketID}},@{name='time_entry_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }    
}

function Update-FreshTicketTimeEntry {
    <#
    .SYNOPSIS
        This function updates a Time Entry.
    .DESCRIPTION
        This function modifies a Time Entry.
    .EXAMPLE
        Update-FreshTicketTimeEntry -TicketID 3001 -agentID 7007 -TimeSpent "01:30"
        Updates a time entry to say agent 7007 spent 1 1/2 hours on the ticket.
    .EXAMPLE
        Update-FreshTicketTimeEntry -TicketID 3002 -agentID 9090 -TimerRunning $false
        Modifies the timer entry and stops the timer.
    .INPUTS
        Object. Either an object representing a Fresh ticket or a time entry.
    .OUTPUTS
        Object. A representation of a time entry.
    #>
    [CmdletBinding()]
    param (
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh Time Entry ID
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('time_entry_id')]        
        [int64]$TimeEntryID,

        # Set to true if timer is currently running. Default value is false. At a time, only one timer can be running for an agent across the account
        [Alias('timer_running')]
        $TimerRunning,

        # Set as true if the time-entry is billable. Default value is true
        $Billable,

        # The total amount of time spent by the timer in hh:mm format. This field cannot be set if timer_running is true. Mandatory if timer_running is false
        [ValidatePattern("^\d{2}:\d{2}$")]
        [Alias('time_spent')]
        [string]$TimeSpent,

        # Time at which the timer is executed. Default value (unless given in request) is the time at which timer is added. Should be less than or equal to current date_time
        [Alias('executed_at')]
        [datetime]$ExecutedAt,

        # Id of the task assigned to the time-entry. Task should be valid on the given ticket and assigned to agent_id
        [Alias('task_id')]
        [int64]$TaskID,

        # Description of the time-entry
        [string]$Note,

        # The user/agent to whom this time-entry is assigned
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias('agent_id')]
        [int64]$AgentID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment         
    )
    
    process {
        if ($null -ne $TimerRunning)
        {
            if ($TimerRunning -isnot [boolean])
            {
                throw 'TimerRunning should be either $true or $false'
            } else {
                Write-Verbose 'TimerRunning is set to a boolean value'
            }
        } else {
            Write-Verbose 'TimerRunning is null'
        }

        if ($null -ne $Billable)
        {
            if ($Billable -isnot [boolean])
            {
                throw 'Billable should be either $true or $false'
            } else {
                Write-Verbose 'Billable is set to a boolean value'
            }
        } else {
            Write-Verbose 'Billable is null'
        }       

        $TimeEntry = @{
            timer_running = $TimerRunning
            billable = $Billable
            time_spent = $TimeSpent
            executed_at = Convert-FreshSafeDate $ExecutedAt
            task_id	 = $TaskID
            note = $Note
            agent_id = $AgentID
        }

        # find empty keys
        $EmptyKeys = @()
        foreach ($key in $TimeEntry.keys)
        {
            if ($null -eq $TimeEntry[$key] -or ($TimeEntry[$key] -isnot [boolean] -and ($TimeEntry[$key] -eq '' -or $TimeEntry[$key] -eq 0)))
            {
                $EmptyKeys += $key
            }
        }

        # remove empty keys
        foreach ($key in $EmptyKeys)
        {
            $TimeEntry.Remove($key)
        }

        $jsonBody = @{time_entry=$TimeEntry} | ConvertTo-Json

        try {
            Invoke-FreshAPIPut -path "tickets/$TicketID/time_entries/$TimeEntryID" -field "time_entry" -body $jsonBody -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='ticket_id';exp={$TicketID}},@{name='time_entry_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Remove-FreshTicketTimeEntry {
    <#
    .SYNOPSIS
        This function deletes an existing Time Entry.
    .DESCRIPTION
        This function deletes an existing Time Entry. Deleted time entries cannot be restored.
    .EXAMPLE
        Get-FreshTicketTimeEntry -TicketID 90210 | Remove-FreshTicketTimeEntry
        Deletes all time entries associated with Tikcet 90210
    .OUTPUTS
        None.
    .INPUTS
        Object. This should either be a Fresh ticket or a time entry object.
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve activites for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh Time Entry ID
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('time_entry_id')]
        $TimeEntryID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    
    process {
        try {
            Invoke-FreshAPIDelete -path "tickets/$TicketID/time_entries/$TimeEntryID" -System $System -verbose:($VerbosePreference -ne 'SilentlyContinue')
        } catch {
            $_ | Convert-FreshError
        }        
    }
    
}

# SERVICE REQUEST RELATED FUNCTIONS

function New-FreshServiceRequest {
    <#
    .SYNOPSIS
        This function creates a service request.
    .DESCRIPTION
        This function creates a service request.
    .NOTES
        Fields behave like the agent portal's new service request page. If a field is not visible in self service portal, you can still provide a value for that field using this function. 
        If a field is marked mandatory, but not visible in portal in service item, you must provide a value for it.
    .INPUTS
        Object. The service cataog item object.
    .OUTPUTS
        Object. A representation of the service request.
    .EXAMPLE
        New-FreshServiceRequest -ItemID 4 -Quantity 1 -Email John.Smith@Who.org
        Creates a new service request for a service catalog item with display_id 4 for John Smith.
    .EXAMPLE
        Get-FreshServiceCatalogItem | Where-Object Name -eq 'Tesla Coils' | New-FreshServiceRequest -Quantity 4 -Email Nikola.Tesla@not-edison.com
        Creates a new service request for 4 Tesla coils       
    #>
    [cmdletBinding()]
    param(
        # Catalogue Item display id
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('display_id','DisplayID','service_item_id')]
        $ItemID,
                
        # Quantity needed by the requested. Must be greater than 0
        [parameter(Mandatory=$true)]
        [ValidateScript({$_ -gt 0})]
        [int]$Quantity,

        # Email id of the requester on whose behalf the service request is created
        [Alias('requested_for')]
        [string]$RequestedForEmail,	

        # Email id of the requester.
        [parameter(Mandatory=$true)]
        [string]$Email,

        # Service items that are included as child items in a hashtable. Provide the display id as service_item_id for each child item e.g. @{service_item_id=22;quantity=1}
        [Alias('child_items')]
        [hashtable]$ChildItems,

        # Values of custom fields present in the service item form. These need to be entered as a hashtable e.g. @{note = "This is a request for a mobile phone";make = "Apple";model="iPhone 13 Pro"}
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]        
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment          
    )

    $Body = @{
        quantity = $Quantity
        requested_for = $RequestedForEmail
        email = $Email
        child_items = $ChildItems
        custom_fields = $CustomFields
    }

    # Find empty keys
    $EmptyKeys = @()
    foreach ($key in $Body.Keys)
    {
        if ($null -eq $Body[$key] -or $Body[$key] -eq '')
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
        Invoke-FreshAPIPost -path "service_catalog/items/$ItemID/place_request" -field "service_request" -body $jsonBody -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='status_text';exp={Get-FreshTicketStatus $_.status}},@{name='priority_text';exp={Get-FreshTicketPriority $_.priority}},@{name='source_text';exp={Get-FreshTicketSource $_.source}},@{name='urgency_text';exp={Get-FreshTicketUrgency $_.urgency}},@{name='impact_text';exp={Get-FreshTicketImpact $_.impact}},@{name='ticket_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }    

}

function Get-FreshRequestedItem {
    <#
    .SYNOPSIS
        View requested items attached to a service request
    .DESCRIPTION
        View requested items attached to a service request
    .EXAMPLE
        Get-FreshRequestedItem -TicketID 555
        Returns all the fields, including the custom fields (custom_fields) for items attached to ticket 555
    .INPUTS
        Object. A representation of the fresh ticket/service request.
    .OUTPUTS
        Object[]. An array of requested item objects.
    .NOTES
        The stage_text field has been added to the returned data to make the data more readable.
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve requested items for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )
    process {
        try {
            Invoke-FreshAPIGet -path "tickets/$TicketID/requested_items" -field "requested_items" -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requested_item_id';exp={$_.id}},@{name='stage_text';exp={Get-FreshSRStage $_.stage}},@{name='ticket_id';exp={$TicketID}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Update-FreshRequestedItem {
    <#
    .SYNOPSIS
        This function updates a service request.
    .DESCRIPTION
        This function updates a service request.
    .INPUTS
        Object. A representation of the updated requested item, with ticket_id & requested_item_id properties.
    .OUTPUTS
        Object. A representation of the updated requested item.
    .EXAMPLE
        Get-FreshRequestedItem -TicketID 123 | Update-FreshRequestedItem -Stage 3 -Remarks "No longer required"
        Cancels all requested items associated with ticket 123
    #>
    [cmdletBinding()]
    param(
        # Fresh Service Request (Ticket) ID
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        $TicketID,

        # ID of Requested Item to update
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('requested_item_id')]
        $RequestedItemID,
                
        # Quantity needed by the requested.
        [int]$Quantity,

        # Email id of the requester on whose behalf the service request is created
        [Alias('requested_for')]
        [string]$RequestedForEmail,	

        # Email id of the requester.
        [string]$Email,

        # Stage of the requested item. Requested = 1, Delivered = 2, Cancelled = 3, Fulfilled = 4, Partially Fulfilled = 5
        [ValidateRange(1,5)]
        [int]$Stage,

        # For cancelled items remarks are mandatory. For other stages, remarks will be ignored.
        [string]$Remarks,

        # Service items that are included as child items in a hashtable. Provide the display id as service_item_id for each child item e.g. @{service_item_id=22;quantity=1}
        [Alias('child_items')]
        [hashtable]$ChildItems,

        # Values of custom fields present in the service item form. These need to be entered as a hashtable e.g. @{note = "This is a request for a mobile phone";make = "Apple";model="iPhone 13 Pro"}
        [Alias('custom_fields')]
        [hashtable]$CustomFields,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment          
    )

    # remarks are only required if item cancelled
    if ($Remarks -ne '' -and $Stage -ne 3)
    {
        Write-Verbose "Remarks are not accepted if the stage is not 3 (Cancelled). Removing remarks..."
        $Remarks = $null
    } elseif ($Stage -eq 3 -and $Remarks -eq '')
    {
        throw "Remarks are MANDATORY when an item is cancelled (Stage 3)"
    }

    $Body = @{
        quantity = $Quantity
        requested_for = $RequestedForEmail
        email = $Email
        stage = $Stage
        remarks = $Remarks
        child_items = $ChildItems
        custom_fields = $CustomFields
    }

    # Find empty keys
    $EmptyKeys = @()
    foreach ($key in $Body.Keys)
    {
        if ($null -eq $Body[$key] -or $Body[$key] -eq '')
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
        Invoke-FreshAPIPut -path "tickets/$TicketID/requested_items/$RequestedItemID" -field "requested_item" -body $jsonBody -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='requested_item_id';exp={$_.id}},@{name='stage_text';exp={$_.stage.name}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }    

}

# TASK RELATED FUNCTIONS

function New-FreshTicketTask {
    <#
    .SYNOPSIS
        Creates a new task in a ticket
    .DESCRIPTION
        Creates a new task in a ticket
    .EXAMPLE
        New-FreshTicketTask -TicketID 12345 -Title "My task" -Description "Check that the other tasks have been completed."
        Creates a task.
    .INPUTS
        Object. the Fresh Ticket to create the task for.
    .OUTPUTS
        Object. The task object.
    #>
    param(
        # Fresh ticket ID to create task for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        $TicketID,

        # Id of the agent to whom the task is assigned
        [Alias('agent_id')]
        [int64]$AgentID,

        # Status of the task, 1-Open, 2-In Progress, 3-Completed
        [ValidateRange(1,3)]
        [int]$Status,
        
        # Due date of the task
        [Alias('due_date')]
        [datetime]$DueDate,

        # Time (in seconds) before which notification is sent prior to due date
        [Alias('notify_before')]
        [int64]$NotifyBefore,

        # Title of the task
        [parameter(Mandatory=$true)]
        [string]$Title,
        
        # Description of the task
        [parameter(Mandatory=$true)]
        [string]$Description,

        # Unique ID of the group to which the task is assigned
        [Alias('group_id')]
        [int64]$GroupID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment         
    )

    $Body = @{
        agent_id = $AgentID
        status = $Status
        due_date = Convert-FreshSafeDate $DueDate
        notify_before = $NotifyBefore
        title = $Title
        description = $Description
        group_id = $GroupID
    }

    # Find empty keys
    $EmptyKeys = @()
    foreach ($key in $Body.Keys)
    {
        if ($null -eq $Body[$key] -or $Body[$key] -eq '')
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
        Invoke-FreshAPIPost -path "tickets/$TicketID/tasks" -field "task" -body $jsonBody -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='ticket_id';exp={$TicketID}},@{name='task_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }        
}

function Get-FreshTicketTask {
    <#
    .SYNOPSIS
        Retrieve all or individual  tasks on a Ticket with the given ID from Freshservice.
    .DESCRIPTION
        Retrieve all or individual  tasks on a Ticket with the given ID from Freshservice.
    .INPUTS
        Object. A task object or a Fresh ticket.
    .OUTPUTS
        Object[]. An array of tasks.
    .EXAMPLE
        Get-FreshTicketTask -TicketID 55
        Returns all tasks associated with ticket #55
    .EXAMPLE
        Get-FreshTicketTask -TicketID 55 -TaskID 41
        Returns the individual task.
    #>
    [CmdletBinding(DefaultParameterSetName='tasks')]
    param(
        # Fresh ticket ID to retrieve requested items for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # ID of task to retrieve from ticket
        [parameter(Mandatory=$true,
            ParameterSetName='task',
            ValueFromPipelineByPropertyName=$true)]
        [Alias('task_id')]
        $TaskID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment            
    )
    process {
        $path = "tickets/$TicketID/tasks"
        $field = $PSCmdlet.ParameterSetName

        if ($TaskID)
        {
            $path += "/$TaskID"
        }
        try {
            Invoke-FreshAPIGet -path $path -field $field -System $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='ticket_id';exp={$TicketID}},@{name='task_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Update-FreshTicketTask {
    <#
    .SYNOPSIS
        Updates a new task in a ticket
    .DESCRIPTION
        Updates a new task in a ticket
    .EXAMPLE
        Update-FreshTicketTask -TicketID 12345 -Status 3
        Marks a task as completed.
    .EXAMPLE
        Get-FreshTicketTask -TicketID 55 | Update-FreshTicketTask -Status 2
        Marks all task on ticket 55 as 'In Progress'
    .INPUTS
        Object. A task.
    .OUTPUTS
        Object. The updated task.
    #>
    param(
        # ID of the Fresh ticket to update task on
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        $TicketID,

        # ID of the task to update
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('task_id')]
        $TaskID,        

        # Id of the agent to whom the task is assigned
        [Alias('agent_id')]
        [int64]$AgentID,

        # Status of the task, 1-Open, 2-In Progress, 3-Completed
        [ValidateRange(1,3)]
        [int]$Status,
        
        # Due date of the task
        [Alias('due_date')]
        [datetime]$DueDate,

        # Time (in seconds) before which notification is sent prior to due date
        [Alias('notify_before')]
        [int64]$NotifyBefore,

        # Title of the task
        [string]$Title,
        
        # Description of the task
        [string]$Description,

        # Unique ID of the group to which the task is assigned
        [Alias('group_id')]
        [int64]$GroupID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment         
    )
    process {
        $Body = @{
            agent_id = $AgentID
            status = $Status
            due_date = Convert-FreshSafeDate $DueDate
            notify_before = $NotifyBefore
            title = $Title
            description = $Description
            group_id = $GroupID
        }

        # Find empty keys
        $EmptyKeys = @()
        foreach ($key in $Body.Keys)
        {
            if ($null -eq $Body[$key] -or $Body[$key] -eq '')
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
            Invoke-FreshAPIPut -path "tickets/$TicketID/tasks/$TaskID" -field "task" -body $jsonBody -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='ticket_id';exp={$TicketID}},@{name='task_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }        
    }
}

function Remove-FreshTicketTask {
    <#
    .SYNOPSIS
        Delete the task on a Ticket.
    .DESCRIPTION
        Delete the task on a Ticket with the given ID from Freshservice
    .INPUTS
        Object. A representation of the Fresh task to delete.
    .OUTPUTS
        None.
    .EXAMPLE
        Get-FreshTicketTask -TicketID 307 | Where-Object Status -eq 3 | Remove-FreshTicketTask
        Deletes all completed tasks from ticket 307
    #>
    [cmdletBinding()]
    param(
        # Fresh ticket ID to delete task from
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # ID of task to retrieve from ticket
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('task_id')]
        $TaskID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment         
    )

    process {
        try {
            Invoke-FreshAPIDelete -path "tickets/$TicketID/tasks/$TaskID" -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue')
        } catch {
            $_ | Convert-FreshError
        }            
    }
}

# CONVERSATION RELATED FUNCTIONS

function New-FreshTicketReply {
    <#
    .SYNOPSIS
        Creates a reply to a ticket
    .DESCRIPTION
        Creates a reply to a ticket
    .INPUTS
        Object. A Fresh ticket object.
    .OUTPUTS
        Object. A Fresh conversation object.
    #>
    [cmdletBinding()]
    param(
            # Fresh ticket ID to retrieve requested items for
            [parameter(Mandatory=$true, 
                ValueFromPipelineByPropertyName=$true)]
            [Alias('ticket_id')]
            [int64]$TicketID,
    
            # HTML body of the reply
            [parameter(Mandatory=$true)]
            [string]$Body,

            # Attachments to add. The total size of all the ticket's attachments (not just this note) cannot exceed 15MB.
            [string[]]$Attachments,

            # The email address from which the reply is sent. By default the global support email will be used.
            [Alias('from_email')]
            [string]$FromEmail,
            
            # ID of the agent/user who is adding the note
            [Alias('user_id')]
            [int64]$UserId,

            # Email address added in the 'cc' field of the outgoing ticket email.
            [Alias('cc_emails')]
            [string[]]$CCEmails,

            # Email address added in the 'bcc' field of the outgoing ticket email.
            [Alias('bcc_emails')]
            $BccEmails,

            # Fresh system/environment to query
            [parameter(ValueFromPipelineByPropertyName=$true)]
            [ValidateSet('Live','Sandbox')]
            [string]$System = $DefaultFreshEnvironment
    )
    $BodyHash = @{
        body = $Body
        'attachments[]' = $Attachments
        from_email = $FromEmail
        user_id	= $UserID
        cc_emails = $CCEmails
        bcc_emails = $BccEmails
    }

    # Find empty keys
    $EmptyKeys = @()
    foreach ($key in $BodyHash.Keys)
    {
        if ($null -eq $BodyHash[$key] -or $BodyHash[$key] -eq '')
        {
            $EmptyKeys += $key
        }
    }

    # remove empty keys
    foreach ($key in $EmptyKeys)
    {
        $BodyHash.Remove($key)
    }

    # If this has attachments, it needs to use multipart/form
    if ($BodyHash.Keys -contains "attachments[]")
    {
        # Includes attachments
        $boundary = [System.Guid]::NewGuid().ToString()
        $ContentType = "multipart/form-data; boundary=`"$boundary`"" 
        $PostBody = ConvertTo-MultiPartFormData -parameters $BodyHash -boundary $boundary
    } else {
        # No attachments - use application/json
        $ContentType = 'application/json'
        $PostBody = $BodyHash | Convertto-Json -depth 5
    }

    try {
        Invoke-FreshAPIPost -path "tickets/$TicketID/reply" -field 'conversation' -body $PostBody -ContentType $ContentType -System $System -Verbose:$verbosity | Select-Object *,@{name='conversation_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }

}

function New-FreshTicketNote {
    <#
    .SYNOPSIS
        Adds a note to a ticket
    .DESCRIPTION
        Adds a note to a ticket
    .INPUTS
        Object. A Fresh ticket object.
    .OUTPUTS
        Object. A conversation object.
    #>
    [cmdletBinding()]
    param(
            # Fresh ticket ID to retrieve requested items for
            [parameter(Mandatory=$true, 
                ValueFromPipelineByPropertyName=$true)]
            [Alias('ticket_id')]
            [int64]$TicketID,
    
            # HTML body of the reply
            [parameter(Mandatory=$true)]
            [string]$Body,

            # Attachments to add. The total size of all the ticket's attachments (not just this note) cannot exceed 15MB.
            [string[]]$Attachments,

            # Set to true if a particular note should appear as being created from the outside (i.e., not through the web portal). The default value is false
            $Incoming,

            # Email addresses of agents/users who need to be notified about this note
            [Alias('notify_emails')]
            [string[]]$NotifyEmails,

            # Set to true if the note is private. The default value is true.
            $Private,

            # ID of the agent/user who is adding the note
            [Alias('user_id')]
            [int64]$UserId,

            # Fresh system/environment to query
            [parameter(ValueFromPipelineByPropertyName=$true)]
            [ValidateSet('Live','Sandbox')]
            [string]$System = $DefaultFreshEnvironment
    )

    process {    
        $BodyHash = @{
            body = $Body
            'attachments[]' = $Attachments
            incoming = $Incoming
            notify_emails = $NotifyEmails
            private = $Private
            user_id	= $UserID
        }

        # Find empty keys
        $EmptyKeys = @()
        foreach ($key in $BodyHash.Keys)
        {
            if ($null -eq $BodyHash[$key] -or ($BodyHash[$key] -isnot [boolean] -and $BodyHash[$key] -eq ''))
            {
                $EmptyKeys += $key
            }
        }

        # remove empty keys
        foreach ($key in $EmptyKeys)
        {
            $BodyHash.Remove($key)
        }

        # If this has attachments, it needs to use multipart/form
        if ($BodyHash.Keys -contains "attachments[]")
        {
            # Includes attachments
            $boundary = [System.Guid]::NewGuid().ToString()
            $ContentType = "multipart/form-data; boundary=`"$boundary`"" 
            $PostBody = ConvertTo-MultiPartFormData -parameters $BodyHash -boundary $boundary
        } else {
            # No attachments - use application/json
            $ContentType = 'application/json'
            $PostBody = $BodyHash | Convertto-Json -depth 5
        }

        try {
            Invoke-FreshAPIPost -path "tickets/$TicketID/reply" -field 'conversation' -body $PostBody -ContentType $ContentType -System $System -Verbose:$verbosity | Select-Object *,@{name='conversation_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Get-FreshTicketConversation {
    <#
    .SYNOPSIS
        View conversations associated to a ticket
    .DESCRIPTION
        View conversations associated to a ticket
    .EXAMPLE
        Get-FreshTicketConversation -TicketID 555
        Returns all conversations associated with ticket 555
    .INPUTS
        Object. A Fresh ticket.
    .OUTPUTS
        Object[]. An array of fresh conversations.
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve conversations for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh system/environment to use
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment           
    )

    process {
        try {
            Invoke-FreshAPIGet -path "tickets/$TicketID/conversations" -field "conversations" -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='conversation_id';exp={$_.id}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

function Update-FreshTicketNote {
    <#
    .SYNOPSIS
        Modifies a note to a ticket
    .DESCRIPTION
        Modifies a note to a ticket. Only public and private notes that haven't been added by the 'System' can be modified.
    .NOTES
        This function returns a 405 error (Method (PUT) not allowed) if you attempt to update a 'System' note.
    #>
    [cmdletBinding()]
    param(
            # ID of the note/conversation to modify
            [parameter(Mandatory=$true,
                ValueFromPipelineByPropertyName=$true)]
            [Alias('conversation_id')]
            [int64]$ConversationID,
    
            # HTML body of the reply
            [parameter(Mandatory=$true)]
            [string]$Body,

            # Attachments to add. The total size of all the ticket's attachments (not just this note) cannot exceed 15MB.
            [string[]]$Attachments,

            # Fresh system/environment to query
            [parameter(ValueFromPipelineByPropertyName=$true)]
            [ValidateSet('Live','Sandbox')]
            [string]$System = $DefaultFreshEnvironment
    )

    $BodyHash = @{
        body = $Body
        'attachments[]' = $Attachments
    }

    # Find empty keys
    $EmptyKeys = @()
    foreach ($key in $BodyHash.Keys)
    {
        if ($null -eq $BodyHash[$key] -or ($BodyHash[$key] -isnot [boolean] -and $BodyHash[$key] -eq ''))
        {
            $EmptyKeys += $key
        }
    }

    # remove empty keys
    foreach ($key in $EmptyKeys)
    {
        $BodyHash.Remove($key)
    }

    # If this has attachments, it needs to use multipart/form
    if ($BodyHash.Keys -contains "attachments[]")
    {
        # Includes attachments
        $boundary = [System.Guid]::NewGuid().ToString()
        $ContentType = "multipart/form-data; boundary=`"$boundary`"" 
        $PostBody = ConvertTo-MultiPartFormData -parameters $BodyHash -boundary $boundary
    } else {
        # No attachments - use application/json
        $ContentType = 'application/json'
        $PostBody = $BodyHash | Convertto-Json -depth 5
    }

    try {
        Invoke-FreshAPIPut -path "conversations/$ConversationID" -field 'conversation' -body $PostBody -ContentType $ContentType -System $System -Verbose:$verbosity | Select-Object *,@{name='conversation_id';exp={$_.id}},@{name='system';exp={$System}}
    } catch {
        $_ | Convert-FreshError
    }
}

function Remove-FreshTicketConversation {
    <#
    .SYNOPSIS
        Delete a conversation on a Ticket.
    .DESCRIPTION
        Delete a conversation on a Ticket with the given ID from Freshservice
    .INPUTS
        Object. A representation of the Fresh conversation to delete.
    .OUTPUTS
        None.
    .EXAMPLE
        Remove-FreshTicketConversation -ConversationID -987
        Deletes this note/reply.
    #>
    [cmdletBinding()]
    param(
        # ID of task to retrieve from ticket
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('conversation_id')]
        [int64]$ConversationID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment         
    )

    process {
        try {
            Invoke-FreshAPIDelete -path "conversations/$ConversationID" -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue')
        } catch {
            $_ | Convert-FreshError
        }            
    }
}

function Remove-FreshTicketConversationAttachment {
    <#
    .SYNOPSIS
        Delete an attachment from a conversation on a Ticket.
    .DESCRIPTION
        Delete an attachment from a conversation on a Ticket with the given ID from Freshservice
    .INPUTS
        Object. A representation of the Fresh attachment to delete.
    .OUTPUTS
        None.
    .EXAMPLE
        Remove-FreshTicketConversationAttachment -ConversationID -987 -AttachmentID 782
        Deletes the attachment from the note.
    #>
    [cmdletBinding()]
    param(
        # ID of task to retrieve from ticket
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('conversation_id')]
        [int64]$ConversationID,

        # ID of the attachment to remove
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('attachment_id')]
        [int64]$AttachmentID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment         
    )

    process {
        try {
            Invoke-FreshAPIDelete -path "conversations/$ConversationID/attachments/$AttachmentID" -system $System -Verbose:($VerbosePreference -ne 'SilentlyContinue')
        } catch {
            $_ | Convert-FreshError
        }            
    }
}

# CSAT Response
function Get-FreshTicketCSATResponse {
    <#
    .SYNOPSIS
        Retrieve a CSAT response of a Ticket with the given ID from Freshservice
    .DESCRIPTION
        Retrieve a CSAT response of a Ticket with the given ID from Freshservice
    .INPUTS
        Object. A Fresh ticket object.
    .OUTPUTS
        Object. The CSAT object contains all the questoinnaire responses.
    #>
    [CmdletBinding()]
    param(
        # Fresh ticket ID to retrieve requested items for
        [parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true)]
        [Alias('ticket_id')]
        [int64]$TicketID,

        # Fresh system/environment to query
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Live','Sandbox')]
        [string]$System = $DefaultFreshEnvironment          
    )

    process {
        try {
            Invoke-FreshAPIGet -path "tickets/$TicketID/csat_response" -field "csat_response" -System $System -Verbose:($VerbosePreference -ne 'SilentlyContinue') | Select-Object *,@{name='ticket_id';exp={$TicketID}},@{name='system';exp={$System}}
        } catch {
            $_ | Convert-FreshError
        }
    }
}

Set-Alias -Name Remove-FreshTicketNote -Value Remove-FreshTicketConversation
Set-Alias -Name Remove-FreshTicketNoteAttachment -Value  Remove-FreshTicketConversationAttachment

Export-ModuleMember -Function Get-FreshRequestedItem,Get-FreshTicket,Get-FreshTicketActivity,Get-FreshTicketConversation,Get-FreshTicketCSATResponse,Get-FreshTicketField,Get-FreshTicketTask,Get-FreshTicketTimeEntry,`
    New-FreshChildTicket,New-FreshServiceRequest,New-FreshTicket,New-FreshTicketNote,New-FreshTicketReply,New-FreshTicketTask,New-FreshTicketTimeEntry,`
    Remove-FreshTicket,Remove-FreshTicketAttachment,Remove-FreshTicketConversation,Remove-FreshTicketConversationAttachment,Remove-FreshTicketTask,Remove-FreshTicketTimeEntry,`
    Restore-FreshTicket,Update-FreshRequestedItem,Update-FreshTicket,Update-FreshTicketNote,Update-FreshTicketTask,Update-FreshTicketTimeEntry `
    -Alias Remove-FreshTicketNote,Remove-FreshTicketNoteAttachment
