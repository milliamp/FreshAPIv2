# Fresh API functions module
# This module provides some PowerShell functions for interacting with Fresh via the Fresh API v2.
# Details of the API can be found at https://api.freshservice.com/
#
# Les Newbigging 2022

# Forcing Tls1.2 to avoid SSL failures
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Retrieving some details from the json files
$Config = Get-Content $PSScriptRoot\Config.json | ConvertFrom-Json

# This variable to be stored globally to be accessible by the nested modules
$DefaultFreshEnvironment = $Config.DefaultEnvironment

# Build authorization headers
$FreshHeaders = @{}
$Environments = "Live","Sandbox"

foreach ($environment in $Environments)
{
    $thisKey = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("$($Config.APIKeys.$environment):X")))
    $FreshHeaders[$environment] = @{Authorization = "Basic $thisKey"}
}

# API Invocation functions
function Invoke-FreshAPIGet {
    <#
    .SYNOPSIS
        Queries the Fresh API
    .DESCRIPTION
        GETs the request from the API endpoint.
    .INPUTS
        None. This does not accept pipeline input.
    .OUTPUTS
        Object[]. An array of fresh response objects.
    .NOTES
        This function is only meant for use within other Fresh functions.
    #>
    [CmdletBinding()]
    param(
        # path of the request 
        [parameter(Mandatory=$true,
                    position=0)]
        $path,

        # the field to read
        [parameter(Mandatory=$true,
                    position=1)]
        $field,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        $System = $DefaultFreshEnvironment
    )

    $uri = "$($Config.BaseURLs.$System)/$path"

    # Invoke-WebRequests usually displays a progress bar. The $ProgressPreference variable determines whether this is displayed (default value = Continue)
    $ProgressPreference = 'SilentlyContinue'  
    try {
        do {
            Write-Verbose "Requesting: $uri "

            # Using Invoke-WebRequest as Invoke-RestMethod does not return link to next page
            $response = Invoke-WebRequest -Headers $FreshHeaders[$System] -Uri $uri -UseBasicParsing
            $objResponse = $response.content | ConvertFrom-Json
            $objResponse.$field
            
            if ($null -eq $response.headers.Link)
            {
                if ($null -eq $objResponse.total)
                {
                    $uri = $null
                } else {
                    # For those queries that do not provide links (e.g. filters for tickets), will cycle through until values reach total
                    $runningTotal = ($objResponse.$field | Measure-Object).count
                    $totalToReach = $objResponse.total
                    Write-Verbose "$totalToReach results expected."
                    write-verbose "Running total: $runningTotal"  
                    $pageNumber = 1
                    while ($runningTotal -lt $totalToReach) 
                    {
                        $pageNumber++
                        $response = Invoke-WebRequest -Headers $FreshHeaders[$System] -Uri "$uri&page=$pageNumber" -UseBasicParsing
                        $objResponse = $response.content | ConvertFrom-Json
                        $objResponse.$field
                        $runningTotal += ($objResponse.$field | Measure-Object).count
                        write-verbose "Running total: $runningTotal"                        
                    } 

                    # Clear $uri
                    $uri = $null
                }
                
            } else {
                $uri = $response.headers.Link.split(';')[0].Replace('<','').Replace('>','')
            }
        } until ($null -eq $uri)
    } catch [System.Net.WebException] {
        # A web error has occurred
        $StatusCode = $_.Exception.Response.StatusCode.Value__
        $Headerdetails = $_.Exception.Response.Headers
        $ThisException = $_.Exception

        switch ($StatusCode)
        {
            400 {
                # Bad Request                
                $StreamReader = [System.IO.StreamReader]::new($thisException.Response.GetResponseStream())
                if ($null -ne $StreamReader)
                {
                    $RawStream = $streamReader.ReadToEnd()
                    $streamReader.Close()
                }
                $OutputObject = $RawStream | ConvertFrom-Json 
                if ($null -eq $OutputObject.errors)
                {
                    throw ($OutputObject | Select-Object @{name='error';exp={400}},@{name='Path&Query';exp={$thisException.Response.ResponseUri.PathAndQuery}},*)
                } else {
                    throw ($OutputObject.errors | Select-Object @{name='error';exp={400}},@{name='Path&Query';exp={$thisException.Response.ResponseUri.PathAndQuery}},@{name='description';exp={$OutputObject.description}},*)
                }
            }

            401 {
                # Authentication Failure
                $OutputObject = [PSCustomObject]@{
                    error = 401
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "The Authorization header is either missing or incorrect"
                }
                throw $OutputObject
            }

            403 {
                # Forbidden
                $OutputObject = [PSCustomObject]@{
                    error = 403
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "Unauthorized (403)."
                }                
                throw $OutputObject
            }

            404 {
                # Not Found
                Write-Verbose "Not Found."
                return $null
            }

            405 {
                # Method not allowed
                $OutputObject = [PSCustomObject]@{
                    error = 405
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "Method not allowed ($($ThisException.Response.Method))."
                }                
                throw $OutputObject
            }

            409 {
                # Inconsistent/Conflicting State
                $OutputObject = [PSCustomObject]@{
                    error = 409
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "The resource that is being created/updated is in an inconsistent or conflicting state."
                }                
                throw $OutputObject                
            }

            429 {
                # Too many requests
                $WaitForSeconds = $Headerdetails['Retry-After']
                Write-Verbose "Waiting for $WaitForSeconds seconds..."
                Start-Sleep -second $WaitForSeconds
                $path = $ThisException.Response.ResponseUri.PathAndQuery.Replace("/api/v2/","") 
                Invoke-FreshAPIQuery -path $path -field $field -system $System                
            }

            500 {
                $OutputObject = [PSCustomObject]@{
                    error = 500
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "A Server Error occurred requesting '$uri'. Please verify the input fields before contacting Fresh."
                }                
                throw $OutputObject                     
            }

            default {
                throw
            }
        }

    } catch {
        throw
    }

    $ProgressPreference = 'Continue'
}

function Invoke-FreshAPIPost {
    <#
    .SYNOPSIS
        Posts to the Fresh API
    .DESCRIPTION
        POSTs a request body to the fresh API endpoint and returns a single record.
    .INPUTS
        None. This does not accept pipeline input.
    .OUTPUTS
        Object. The newly created Fresh object
    .NOTES
        This function is only meant to be used by other Fresh functions.
    #>
    [CmdletBinding()]
    param(
        # path of the request
        [parameter(Mandatory=$true)]
        $path,

        # the field name returned
        [parameter(Mandatory=$false)]
        $field,

        # the  body of the request. This can be either json or a formatted form-data
        [parameter(Mandatory=$false)]
        [Alias('json')]
        $body,

        # ContentType. By default is 'application/json'
        $ContentType='application/json',

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        $System = $DefaultFreshEnvironment
    )

    $uri = "$($Config.BaseURLs.$System)/$path"

    try {
        Write-Verbose "Posting to: $uri"
        Write-Verbose "Body:`r`n$body"
        $response = Invoke-RestMethod -Headers $FreshHeaders[$System] -Uri $uri -UseBasicParsing -Method Post -Body $body -ContentType $ContentType
        $response.$field
    } catch [System.Net.WebException] {
        # A web error has occurred
        $StatusCode = $_.Exception.Response.StatusCode.Value__
        $Headerdetails = $_.Exception.Response.Headers
        $ThisException = $_.Exception

        switch ($StatusCode)
        {
            400 {
                # Bad Request                
                $StreamReader = [System.IO.StreamReader]::new($thisException.Response.GetResponseStream())
                if ($null -ne $StreamReader)
                {
                    $RawStream = $streamReader.ReadToEnd()
                    $streamReader.Close()
                }
                $OutputObject = $RawStream | ConvertFrom-Json 
                if ($null -eq $OutputObject.errors)
                {
                    throw ($OutputObject | Select-Object @{name='error';exp={400}},@{name='Path&Query';exp={$thisException.Response.ResponseUri.PathAndQuery}},*)
                } else {
                    throw ($OutputObject.errors | Select-Object @{name='error';exp={400}},@{name='Path&Query';exp={$thisException.Response.ResponseUri.PathAndQuery}},@{name='description';exp={$OutputObject.description}},*)
                }
            }

            401 {
                # Authentication Failure
                $OutputObject = [PSCustomObject]@{
                    error = 401
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "The Authorization header is either missing or incorrect"
                }
                throw $OutputObject
            }

            403 {
                # Forbidden
                $OutputObject = [PSCustomObject]@{
                    error = 403
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "Unauthorized (403)."
                }                
                throw $OutputObject
            }

            404 {
                # Not Found
                Write-Verbose "Not Found."
                return $null
            }

            405 {
                # Method not allowed
                $OutputObject = [PSCustomObject]@{
                    error = 405
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "Method not allowed ($($ThisException.Response.Method))."
                }                
                throw $OutputObject
            }

            409 {
                # Inconsistent/Conflicting State
                $OutputObject = [PSCustomObject]@{
                    error = 409
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "The resource that is being created/updated is in an inconsistent or conflicting state."
                }                
                throw $OutputObject                
            }

            429 {
                # Too many requests
                $WaitForSeconds = $Headerdetails['Retry-After']
                Write-Verbose "Waiting for $WaitForSeconds seconds..."
                Start-Sleep -second $WaitForSeconds
                $path = $ThisException.Response.ResponseUri.PathAndQuery.Replace("/api/v2/","") 
                Invoke-FreshAPIQuery -path $path -field $field -system $System                
            }

            500 {
                $OutputObject = [PSCustomObject]@{
                    error = 500
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "A Server Error occurred requesting '$uri'. Please verify the input fields before contacting Fresh."
                }                
                throw $OutputObject                     
            }

            default {
                throw
            }
        }

    } catch {
        throw
    }
}

function Invoke-FreshAPIPut {
    <#
    .SYNOPSIS
        PUTs data to the Fresh API
    .DESCRIPTION
        PUTs data to the Fresh API endpoint for updating and returns the updated record.
    .INPUTS
        None. This does not accept pipeline input.
    .OUTPUTS
        Object. The updated Fresh object.
    .NOTES
        This function is meant for use by other Fresh functions.
    #>
    [CmdletBinding()]
    param(
        # path of the request
        [parameter(Mandatory=$true)]
        $path,

        # The field to return
        $field,

        # the  body of the request. This can be either json or a formatted form-data
        [parameter(Mandatory=$false)]
        [Alias('json')]
        $body,

        # ContentType. By default is 'application/json'
        $ContentType='application/json',

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        $System = $DefaultFreshEnvironment        
    )

    $uri = "$($Config.BaseURLs.$System)/$path"

    try {
        Write-Verbose "Putting to: $uri"
        if ($body)
        {
            Write-Verbose "Body:`r`n$body"
        }
        $response = Invoke-RestMethod -Headers $FreshHeaders[$System] -Uri $uri -UseBasicParsing -Method Put -Body $body -ContentType $ContentType
        $response.$field
    } catch [System.Net.WebException] {
        # A web error has occurred
        $StatusCode = $_.Exception.Response.StatusCode.Value__
        $Headerdetails = $_.Exception.Response.Headers
        $ThisException = $_.Exception

        switch ($StatusCode)
        {
            400 {
                # Bad Request                
                $StreamReader = [System.IO.StreamReader]::new($thisException.Response.GetResponseStream())
                if ($null -ne $StreamReader)
                {
                    $RawStream = $streamReader.ReadToEnd()
                    $streamReader.Close()
                }
                $OutputObject = $RawStream | ConvertFrom-Json 
                if ($null -eq $OutputObject.errors)
                {
                    throw ($OutputObject | Select-Object @{name='error';exp={400}},@{name='Path&Query';exp={$thisException.Response.ResponseUri.PathAndQuery}},*)
                } else {
                    throw ($OutputObject.errors | Select-Object @{name='error';exp={400}},@{name='Path&Query';exp={$thisException.Response.ResponseUri.PathAndQuery}},@{name='description';exp={$OutputObject.description}},*)
                }
            }

            401 {
                # Authentication Failure
                $OutputObject = [PSCustomObject]@{
                    error = 401
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "The Authorization header is either missing or incorrect"
                }
                throw $OutputObject
            }

            403 {
                # Forbidden
                $OutputObject = [PSCustomObject]@{
                    error = 403
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "Unauthorized (403)."
                }                
                throw $OutputObject
            }

            404 {
                # Not Found
                Write-Verbose "Not Found."
                return $null
            }

            405 {
                # Method not allowed
                $OutputObject = [PSCustomObject]@{
                    error = 405
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "Method not allowed ($($ThisException.Response.Method))."
                }                
                throw $OutputObject
            }

            409 {
                # Inconsistent/Conflicting State
                $OutputObject = [PSCustomObject]@{
                    error = 409
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "The resource that is being created/updated is in an inconsistent or conflicting state."
                }                
                throw $OutputObject                
            }

            429 {
                # Too many requests
                $WaitForSeconds = $Headerdetails['Retry-After']
                Write-Verbose "Waiting for $WaitForSeconds seconds..."
                Start-Sleep -second $WaitForSeconds
                $path = $ThisException.Response.ResponseUri.PathAndQuery.Replace("/api/v2/","") 
                Invoke-FreshAPIQuery -path $path -field $field -system $System                
            }

            500 {
                $OutputObject = [PSCustomObject]@{
                    error = 500
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "A Server Error occurred requesting '$uri'. Please verify the input fields before contacting Fresh."
                }                
                throw $OutputObject                     
            }

            default {
                throw
            }
        }

    } catch {
        throw
    }
}

function Invoke-FreshAPIDelete {
    <#
    .SYNOPSIS
        Requests DELETE of object to the Fresh API
    .DESCRIPTION
        Issues a DELETE to the Fresh API endpoint
    .INPUTS
        None. THis does not accept pipeline input.
    .OUTPUTS
        None (usually).
    .NOTES
        The -field otion is not included because most DELETEs will result in no content; will rely on other functions to handle any returned data appropriately.
        This function is meant to be used only by other Fresh functions.
    #>
    [CmdletBinding()]
    param(
        # path of the request
        [parameter(Mandatory=$true)]
        $path,

        # Fresh system/environment to use
        [ValidateSet('Live','Sandbox')]
        $System = $DefaultFreshEnvironment   
    )

    $uri = "$($Config.BaseURLs.$System)/$path"

    try {
        Write-Verbose "Deleting: $uri"
        Invoke-RestMethod -Headers $FreshHeaders[$System] -Uri $uri -UseBasicParsing -Method Delete
    } catch [System.Net.WebException] {
        # A web error has occurred
        $StatusCode = $_.Exception.Response.StatusCode.Value__
        $Headerdetails = $_.Exception.Response.Headers
        $ThisException = $_.Exception

        switch ($StatusCode)
        {
            400 {
                # Bad Request                
                $StreamReader = [System.IO.StreamReader]::new($thisException.Response.GetResponseStream())
                if ($null -ne $StreamReader)
                {
                    $RawStream = $streamReader.ReadToEnd()
                    $streamReader.Close()
                }
                $OutputObject = $RawStream | ConvertFrom-Json 
                if ($null -eq $OutputObject.errors)
                {
                    throw ($OutputObject | Select-Object @{name='error';exp={400}},@{name='Path&Query';exp={$thisException.Response.ResponseUri.PathAndQuery}},*)
                } else {
                    throw ($OutputObject.errors | Select-Object @{name='error';exp={400}},@{name='Path&Query';exp={$thisException.Response.ResponseUri.PathAndQuery}},@{name='description';exp={$OutputObject.description}},*)
                }
            }

            401 {
                # Authentication Failure
                $OutputObject = [PSCustomObject]@{
                    error = 401
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "The Authorization header is either missing or incorrect"
                }
                throw $OutputObject
            }

            403 {
                # Forbidden
                $OutputObject = [PSCustomObject]@{
                    error = 403
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "Unauthorized (403)."
                }                
                throw $OutputObject
            }

            404 {
                # Not Found
                Write-Verbose "Not Found."
                return $null
            }

            405 {
                # Method not allowed
                $OutputObject = [PSCustomObject]@{
                    error = 405
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "Method not allowed ($($ThisException.Response.Method))."
                }                
                throw $OutputObject
            }

            409 {
                # Inconsistent/Conflicting State
                $OutputObject = [PSCustomObject]@{
                    error = 409
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "The resource that is being created/updated is in an inconsistent or conflicting state."
                }                
                throw $OutputObject                
            }

            429 {
                # Too many requests
                $WaitForSeconds = $Headerdetails['Retry-After']
                Write-Verbose "Waiting for $WaitForSeconds seconds..."
                Start-Sleep -second $WaitForSeconds
                $path = $ThisException.Response.ResponseUri.PathAndQuery.Replace("/api/v2/","") 
                Invoke-FreshAPIQuery -path $path -field $field -system $System                
            }

            500 {
                $OutputObject = [PSCustomObject]@{
                    error = 500
                    'Path&Query' = $thisException.Response.ResponseUri.PathAndQuery
                    description = "A Server Error occurred requesting '$uri'. Please verify the input fields before contacting Fresh."
                }                
                throw $OutputObject                     
            }

            default {
                throw
            }
        }

    } catch {
        throw
    }
}

# Supporting functions required by all modules
function ConvertTo-MultiPartFormData {
    <#
    .SYNOPSIS
        Converts data to multipart form data for POSTing to the API.
    .DESCRIPTION
        Converts data to multipart form data for POSTing to the API.
    .NOTES
        This function is only meant to be used by other Fresh functions.
    #>
    [CmdletBinding()]
    param (
        [hashtable]$parameters,

        [string]$boundary
    )

    process {
        # Create an empty array to start building parameters
        $BodyArray = @()
        $CRLF = "`r`n"

        foreach ($key in $parameters.keys)
        {
            if ($key -eq "attachments[]")
            {
                foreach ($attachment in $parameters[$key])
                {
                    $FileItem = Get-Item $attachment
                    $FileName = $FileItem.Name
                    $FileContents = [System.IO.File]::ReadAlltext($FileItem.FullName)
                    $BodyArray += "--$boundary","Content-Disposition: form-data; name=`"attachments[]`"; filename=`"$FileName`"","Content-Type: application/octet-stream$CRLF",$FileContents
                }
                
            } else {
                if (($parameters[$key] | Measure-Object).count -eq 1)
                {
                    $BodyArray += "--$boundary","Content-Disposition: form-data; name=`"$key`"","Content-Type: text/plain$CRLF",$parameters[$key].ToString()
                } else {
                    foreach ($item in $parameters[$key])
                    {
                        $BodyArray += "--$boundary","Content-Disposition: form-data; name=`"$key`"","Content-Type: text/plain$CRLF",$item.ToString()
                    }
                }
            }
        }

        $BodyArray += "--$boundary--$CRLF" 

        return ($BodyArray -join $CRLF)
        
    }
}

function ConvertTo-ISO8601Date {
    <#
    .SYNOPSIS
        Converts [datetime] objects to ISO8601 format
    .DESCRIPTION
        Converts [datetime] objects to ISO8601 format
    #>
    [cmdletBinding()]
    param(
        # Date to convert
        [datetime]$DateTime,

        # if the time details are not required (date only)
        [switch]$DateOnly
    )
    $dateFormat = 'yyyy-MM-ddTHH:mm:ssK'
    if ($DateOnly)
    {
        $dateFormat = 'yyyy-MM-dd'
    }
    $DateTime.ToUniversalTime().ToString($dateFormat)
}

function Convert-FreshSafeDate {
    <#
    .SYNOPSIS
        This function safely handles input date data.
    .DESCRIPTION
        This will handle dates where a null was passed, and returns a null.
    .NOTES
        This function is meant for use by other Fresh functions only.
    #>
    [cmdletBinding()]
    param(
        $date
    )

    if ($null -ne $date)
    {
        ConvertTo-ISO8601Date $date
    } else {
        return $date
    }
}

function Convert-FreshError {
    <#
    .SYNOPSIS
        Converts returned error from Fresh into something easier to read.
    .DESCRIPTION
        Converts returned error from Fresh into something easier to read.
    #>
    [cmdletBinding()]
    param(
        [parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        $thisError
    )

    foreach ($TargetObject in $thisError.TargetObject)
    {
        if ($TargetObject.description -eq 'Validation failed') # -and $TargetObject.code -eq 'invalid_value')
        {
            switch ($TargetObject.code)
            {
                'invalid_value' {
                    Write-Warning "An invalid value was given for $($TargetObject.field). $($TargetObject.message)"
                }

                'missing_field' {
                    Write-Warning "The $($TargetObject.field) parameter is missing. $($TargetObject.message)"
                }

                'datatype_mismatch' {
                    Write-Warning "There is a data type mismatch for $($TargetObject.field) parameter is missing. $($TargetObject.message)"
                }
            }
            
        } 
    }
    throw $thisError.TargetObject
}

Set-Alias -Name Invoke-FreshAPIQuery -Value Invoke-FreshAPIGet

Export-ModuleMember -Variable DefaultFreshEnvironment -Function * -Alias Invoke-FreshAPIQuery