# FreshAPIv2 PowerShell Module 

This PowerShell module presents a range of functions that work with the Fresh Service APIs to manage ticket, requesters, agents, etc. within the Fresh Service platform.
Details about the Fresh Servcie API can be found at [https://api.freshservice.com](https://api.freshservice.com).

## Setting up the module

After copying the module, you need to configure the API details before you can use this module.Copy (or rename) the config-template.json to config.json. Edit the config.json file, with the following settings:
* BaseURLs: these should be the hostnames used to access your APIs, and is usually the tenant name followed by ".freshservice.com." Do not use any alias for this - it must be the ".freshservice.com" address.. 
* APIKeys: details on how to find your API key can be found on the Fresh Service support site: [Where do I find my API key?](https://support.freshservice.com/en/support/solutions/articles/50000000306-where-do-i-find-my-api-key-) Note that the API will only grant the same access levels and permissions as the account it was obtained from.
* DefaultEnvironment: This is the default system used when functions are called without specifying the -System parameter. If you do not have a sandbox environment, this should be 'Live'. 

## Using the module

Currently (August 2022), the functions available cover most aspects of ticket and requester management. The functionality does not currently include the ability to query with dates (e.g. list all tickets due before 12pm tomorrow), but that will be added a future date.

## Modules contained

To make the modules more manageable, the functions have been broken down into nested modules. The modules are:
* **FreshAPIv2.psm1** - this contains the functions that make the calls to the APIs, as well as other shared functions to support functions in other modules.
* **FreshAPI.Requesters.psm1** - this contains functions relating to Requester, Agent, Requester Group and Agent Group.
* **FreshAPI.Tickets.psm1** - this contains functions relating to tickets.
* **FreshAPI.ServiceCatalog.psm1** - this contains functions relating to the Service Catalog and its categories.
* **FreshAPI.Locations.psm1** - this contains functions relating to Locations in Fresh.
* **FreshAPI.Departments.psm1** - this contains functions relating to departments in Fresh
* **FreshAPI.CustomObjects.psm1** - this contains functions relating to Fresh custom objects and their records

Further supporting modules will be added in due time to cover off the other API areas, as per [Fresh's API documentation](https://api.freshservice.com/)

## Available Functions

Here is a list of the functions available for use. Other functions are exported, but these are supporting functions - the ones below are those intended for use:

### Agent related functions

* New-FreshAgent
* Get-FreshAgent
* Update-FreshAgent
* Remove-FreshAgent
* Restore-FreshAgent
* ConvertTo-FreshRequester
* Get-FreshAgentField

### Agent Group & Role related functions

* New-FreshAgentGroup
* Get-FreshAgentGroup
* Update-FreshAgentGroup
* Remove-FreshAgentGroup
* Get-FreshAgentRole

### Department related functions 

* New-FreshDepartment
* Get-FreshDepartment
* Update-FreshDepartment
* Remove-FreshDepartment
* Get-FreshDepartmentFields

### Location related functions

* New-FreshLocation
* Get-FreshLocation
* Update-FreshLocation
* Remove-FreshLocation

### Requester related functions

* New-FreshRequester
* Get-FreshRequester
* Update-FreshRequester
* Merge-FreshRequester
* Remove-FreshRequester
* Restore-FreshRequester
* ConvertTo-FreshAgent

### Requester group related functions

* New-FreshRequesterGroup
* Get-FreshRequesterGroup
* Update-FreshRequesterGroup
* Remove-FreshRequesterGroup
* Get-FreshRequesterGroupMember
* Add-FreshRequesterGroupMember
* Remove-FreshRequesterGroupMember

### Service Catalog related functions

* Get-FreshServiceCatalogItem
* Search-FreshServiceCatalogItem
* Get-FreshServiceCategory

### Fresh Ticket related functions

* New-FreshTicket
* New-FreshChildTicket
* Get-FreshTicket
* Update-FreshTicket
* Remove-FreshTicket
* Restore-FreshTicket
* Get-FreshTicketActivity
* Remove-FreshTicketAttachment
* Get-FreshTicketField
* Get-FreshTicketCSATResponse

### Service Request related functions

* New-FreshServiceRequest
* Get-FreshRequestedItem
* Update-FreshRequestedItem

### Fresh ticket conversation related functions

* New-FreshTicketNote
* New-FreshTicketReply
* Get-FreshTicketConversation
* Update-FreshTicketNote
* Remove-FreshTicketConversation
* Remove-FreshTicketConversationAttachment

### Task related related functions

* New-FreshTicketTask
* Get-FreshTicketTask
* Update-FreshTicketTask
* Remove-FreshTicketTask

### Time Entry related functions

* New-FreshTicketTimeEntry
* Get-FreshTicketTimeEntry
* Update-FreshTicketTimeEntry
* Remove-FreshTicketTimeEntry

### Asset related functions

* Get-FreshAsset
* Get-FreshAssetType

### Fresh Custom Object related functions

* Get-FreshCustomObject
* New-FreshCustomObjectRecord
* Get-FreshCustomObjectRecord
* Update-FreshCustomObjectRecord
* Remove-FreshCustomObjectRecord
