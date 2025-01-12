<# 
.SYNOPSIS
Ingests webhook information sent from Freshservice Workflow Automator and creates a reservable resource in Microsoft Exchange given information provided.

.DESCRIPTION 
The script is intended to ingest webhook data sent from a defined Freshservice Workflow in the context of an Azure Runbook. 
In this case, a worklfow is defined in my Freshservice tenant to trigger when a requestor requests a published service catalog item. 

The workflow subsequently pulls the information provided from the service request form and submits JSON data to a defined Webhook URL 
associated with the Azure Runbook. Once the webhook data is received, the Azure Runbook will then execute this script and ingest data from
the webhook data and create an Exchange Workspace Resource.

.COMPONENT 
Requires the installation or import of the ExchangeOnlineManagement PowerShell module.

.PARAMETER WebhookData
Ingests expected webhook data from Freshservice to convert and assign to PowerShell objects.
#>

[CmdletBinding(DefaultParameterSetName = "WebhookOnly")]
param (
    # Parameter for only webhook data
    [Parameter(ParameterSetName = "WebhookTrigger", Mandatory = $true)]
    [Object]$WebhookData,

    # Parameter for only webhook data
    [Parameter(ParameterSetName = "JsonOnly", Mandatory = $true)]
    [Object]$AsJSON,

    # Parameter set for Office type.
    [Parameter(ParameterSetName = "OfficeSet", Mandatory = $true)]
    [String]$Office_Of,

    [Parameter(ParameterSetName = "OfficeSet")]
    [String]$OfficeId,

    [Parameter(ParameterSetName = "OfficeSet")]
    [Int]$FloorNum_Of,

    [Parameter(ParameterSetName = "OfficeSet")]
    [String]$FloorLabel_Of,

    [Parameter(ParameterSetName = "OfficeSet")]
    [Int]$Capacity_Of,

    [Parameter(ParameterSetName = "OfficeSet")]
    [Switch]$WheelChairAccessible_Of,

    [Parameter(ParameterSetName = "OfficeSet")]
    [Array]$Delegates_Of,

    [Parameter(ParameterSetName = "OfficeSet")]
    [Array]$CalendarPermissions_Of,

    # Parameter set for Cubicle type.
    [Parameter(ParameterSetName = "CubicleSet", Mandatory = $true)]
    [String]$Office_Cu,

    [Parameter(ParameterSetName = "CubicleSet")]
    [String]$CubicleId,

    [Parameter(ParameterSetName = "CubicleSet")]
    [Int]$FloorNum_Cu,

    [Parameter(ParameterSetName = "CubicleSet")]
    [String]$FloorLabel_Cu,

    [Parameter(ParameterSetName = "CubicleSet")]
    [Int]$Capacity_Cu,

    [Parameter(ParameterSetName = "CubicleSet")]
    [Switch]$WheelChairAccessible_Cu,

    [Parameter(ParameterSetName = "CubicleSet")]
    [Array]$Delegates_Cu,

    [Parameter(ParameterSetName = "CubicleSet")]
    [Array]$CalendarPermissions_Cu
)

$JsonSchema = '{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "type": "object",
  "properties": {
    "Office": {
      "type": "string",
      "description": "The name of the office."
    },
    "FloorNum": {
      "type": "integer",
      "description": "The floor number of the office."
    },
    "FloorLabel": {
      "type": "string",
      "description": "The label for the floor."
    },
    "Capacity": {
      "type": "integer",
      "description": "The capacity of the office."
    },
    "WheelChairAccessible": {
      "type": "boolean",
      "description": "Indicates whether the office is wheelchair accessible."
    },
    "OfficeId": {
      "type": "string",
      "description": "The unique identifier for the office."
    },
    "Delegates": {
      "type": "string",
      "description": "The delegate(s) associated with the office."
    },
    "CalendarPermissions": {
      "type": "array",
      "items": {
        "type": "string"
      },
      "description": "Permissions associated with the office calendar."
    },
    "CubicleId": {
      "type": "string",
      "description": "The unique identifier for the cubicle."
    },
    "TicketId": {
      "type": "integer",
      "description": "The identifier for a related ticket."
    },
    "ServiceRequestItemId": {
      "type": "integer",
      "description": "The identifier for a related service request item."
    }
  },
  "required": ["Office", "FloorNum", "FloorLabel", "Capacity", "WheelChairAccessible", "Delegates", "CalendarPermissions"],
  "oneOf": [
    {
      "required": ["OfficeId"],
      "not": {
        "required": ["CubicleId"]
      }
    },
    {
      "required": ["CubicleId"],
      "not": {
        "required": ["OfficeId"]
      }
    }
  ]
}'

# Update variable names depending on chosen parameter set.
if ($OfficeSet)
{
  $Office = $Office_Of
  $FloorNum = $FloorNum_Of
  $FloorLabel = $FloorLabel_Of
  $Capacity = $Capacity_Of
  $WheelChairAccessible = $WheelChairAccessible_Of
  $OfficeId = $OfficeId
  $Delegates = $Delegates_Of
  $CalendarPermissions = $CalendarPermissions_Of
}
elseif ($CubicleSet)
{
  $Office = $Office_Cu
  $FloorNum = $FloorNum_Cu
  $FloorLabel = $FloorLabel_Cu
  $Capacity = $Capacity_Cu
  $WheelChairAccessible = $WheelChairAccessible_Cu
  $CubicleId = $CubicleId
  $Delegates = $Delegates_Cu
  $CalendarPermissions = $CalendarPermissions_Cu
}
# Populate variables with webhook data, if provided.
elseif ($WebhookData)
{
  # Outputs request header details.
  Write-Output $WebhookData.RequestHeader

  if ($WebhookData.RequestBody) 
  {
    try {
      if ($ValidWebhookData = Test-Json -Json $WebbookData.RequestBody -Schema $JsonSchema)
      {
        # Converts request body from JSON request body to PS Object.
        $PayloadRequestBody = (ConvertFrom-Json -InputObject $WebhookData.RequestBody)
  
        # Set PS variables for basic office attributes.
        $Office = $PayloadRequestBody.location.Trim()
        $FloorNum = $PayloadRequestBody.floornum.Trim()
        $FloorLabel = $PayloadRequestBody.floorlabel.Trim()
        $Capacity = $PayloadRequestBody.capacity.Trim()
        $WheelChairAccessible = $PayloadRequestBody.wheelchairaccessible
        $OfficeId = $PayloadRequestBody.officeid.Trim()
        $CubicleId = $PayloadRequestBody.cubicleid.Trim()
        $Delegates = $PayloadRequestBody.delegates.Split(',').Trim()
        $CalendarPermissions = $PayloadRequestBody.calendarpermissions.Split(',').Trim()
        $TicketID = $PayloadRequestBody.ticketid.Trim()
        $ServiceRequestItemID = ($PayloadRequestBody.itemrequestid -replace '[\[\]]', '').Trim()
      }
    }
    catch {
      Write-Error -Message "$($Error[0].Exception.Message)"
      Exit 1
    }
  }
}
elseif ($JsonOnly)
{  
  if ($ValidJson = Test-Json -Json $Json -Schema $JsonSchema)
  {
    try {
      # Converts request body from JSON request body to PS Object.
      $PSObject = (ConvertFrom-Json -InputObject $Json)

      # Set PS variables for basic office attributes.
      $Office = ($PSObject.location).Trim()
      $FloorNum = ($PSObject.floornum).Trim()
      $FloorLabel = ($PSObject.floorlabel).Trim()
      $Capacity = ($PSObject.capacity).Trim()
      $WheelChairAccessible = $PSObject.wheelchairaccessible
      $OfficeId = ($PSObject.officeid).Trim()
      $CubicleId = ($PSObject.cubicleid).Trim()
      $Delegates = ($PSObject.delegates.Split(',')).Trim()
      $CalendarPermissions = ($PSObject.calendarpermissions).Trim().Split(',')

    }
    catch {
      Write-Error -Message "$($Error[0].Exception.Message)"
      Exit 1
    }
  }
}

# Import Exchange PowerShell module to session.
Import-Module ExchangeOnlineManagement

# [REQUIRED] Set organization name, domain, and Azure subscription ID if using managed identity.
$OrganizationName = 'ORGANIZATION-NAME-HERE'
$FSDomain = 'FRESH-SERVICE-DOMAIN-HERE'
$DomainName = 'M365-DEFAULT-DOMAIN-HERE'
$SubscriptionId = 'AZURE-SUBSCRIPTION-ID-HERE'
$AdminGroup = 'MAIL-ENABLED-SECURITY-GROUP-HERE'

# [REQUIRED] Set office and cubicle abbreviation values as well as resource type to be created.
$OfficeAbbr = 'OF'
$CubicleAbbr = 'WS'
$ResourceType = 'Workspace'

# [OPTIONAL] Set Keyvault and credential name variables for retrieving credentials from Azure Keyvault to authenticate to other API supported systems such as an ITSM.
$KeyvaultName = 'KEY-VAULT-NAME-HERE'
$CredentialName = 'CREDENTIAL-NAME-HERE'


# [OPTIONAL] Set API URLs for Freshservice tenant with unique ticket ID and service request item ID.

# Private note URL with unique ticket ID.
$FreshserviceCreatePrivateNoteUpdateURL = "https://$FSDomain/api/v2/tickets/$TicketID/notes"

# Service Request item URL with unique request item ID
$FreshserviceUpdateServiceRequestItemStatusURL = "https://$FSDomain/api/v2/tickets/$TicketID/requested_items/$ServiceRequestItemID"

# [REQUIRED] Office locations
$OfficeLocations = @(

# Office 1 details
  @{
    Name = 'Main Campus'
    RoomList = 'roomlist1@contoso.com'
    Building = 'Bldg 1'
    Street = '2 Microsoft Way'
    City = 'Redmond'
    State = 'WA'
    Zipcode = '13464'
    Country = 'United States'
  }

  # Office 2 details
  @{
    Name = 'Engineering Bldg'
    RoomList = 'roomlist2@contoso.com'
    Building = 'Bldg 2'
    Street = '2 Microsoft Way'
    City = 'Redmond'
    State = 'WA'
    Zipcode = '13464'
    Country = 'United States'
  }
  # Office 3 details
  @{
    Name = 'Research & Development Bldg'
    RoomList = 'roomlist3@contoso.com'
    Building = 'Bldg 3'
    Street = '3 Microsoft Way'
    City = 'Redmond'
    State = 'WA'
    Zipcode = '13464'
    Country = 'United States'
  }
)

if ($KeyvaultName -and $CredentialName)
{
  # Set parameters for Get-AzKeyVaultSecret cmdlet to securely retrieve Freshservice Agent API creds for Freshservice API requests.
  $KeyVaultParams = @{
    Name = $CredentialName
    VaultName = $KeyvaultName
    AsPlainText = $true
  }
}
if ($ValidWebhookData)
{
  try {
  # Connect to Azure for retrieving credentials.
  Connect-AzAccount -Subscription $SubscriptionId -Identity
    
  }
  catch {
    Write-Error -Message "$($Error[0].Exception.Message)"
    Exit 1
  }

  # Sets header info for Freshservice API call. Retrieves Freshservice API key from Azure Key Vault and encodes using Base64.
  $Headers = @{
    "Authorization" = ("Basic" + " " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('{0}:{1}' -f (Get-AzKeyVaultSecret @KeyVaultParams), $null))) )
    "Content-Type" = "application/json"
  }
}
  # Sets room list info based on provided office name.
  foreach ($Location in $OfficeLocations)
  {
    if ($Office_Of -eq $Location.Name)
    {
      $Office = $Location
    }
    elseif ($Office_Cu -eq $Location.Name)
    {
      $Office = $Location
    }
  }

    # Set username using office abbreviation code and office ID #.
    if ($OfficeId)
    {
      $IsOffice = $true
      $Username = ($Office.Bldg.ToLower() + "-" + $OfficeAbbr.ToLower() + "-" + $OfficeId)
      $DisplayName = ($OfficeAbbr + " " + $OfficeId.ToUpper())
    }
    if ($CubicleId)
    {
      $Username = ($Office.Bldg.ToLower() + "-" + $CubicleAbbr.ToLower() + "-" + $CubicleId)
      $DisplayName = ($CubicleAbbr + " " + $CubicleId.ToUpper())
    }

    ################################################################################################################  **Parameters for resource settings**  ###################################################################################################################################################

    # Base parameters for Set-Place cmdlet. Indicates basic resource location information.
    $SetPlaceParams = @{
      Identity = $Username
      Building = $Office
      Capacity = $Capacity
      Street = $Street
      City = $City
      State = $State
      PostalCode = $Zipcode
      CountryOrRegion = $Country
      Floor = $FloorNum
      FloorLabel = $FloorLabel
    }

    # Appends parameter to Set-Place cmdlet if resource is handicap accessible.
    if ($WheelChairAccessible)
    {
      # Parameters for Set-Place cmdlet.
      $SetPlaceParams += @{ IsWheelChairAccessible = $true }
    }

    # Parameters for Set-CalendarProcessing cmdlet.
    $SetCalendarProcessingParams = @{
      Identity = $Username
      AutomateProcessing = "AutoAccept"
      AllowConflicts = $false
      AllowRecurringMeetings = $true
      EnforceCapacity = $true
      RemoveOldMeetingMessages = $true
      RemoveCanceledMeetings = $true
      Confirm = $false
    }

    # Parameters for Add-DistributionGroupMember cmdlet. Adds resource to room list (distribution list group) to allow resource to be found in Outlook Room Finder tool.
    $AddDistributionGroupMemberParams = @{
      Identity = $RoomList
      Member = $Username
      Confirm = $false
    }

    # Parameters for New-Mailbox cmdlet. Creates room resource.
    $NewMailboxParams = @{
      Name = $Username
      Room = $true
      Confirm = $false
    }

    # Parameters for Add-MailboxFolderPermission cmdlet. Adds editor rights to administrative groups.
    $AddMailboxParams = @{
      Identity = $Username + ":\calendar"
      AccessRights = "Editor"
      Confirm = $false
    }

    # Parameters for Set-Mailbox cmdlet. Sets Display Name, Name, and type to Workspace.
    $SetMailboxParams = @{
      Identity = $Username
      Type = $ResourceType
      Name = $DisplayName
      DisplayName = $DisplayName
      Confirm = $false
    }

    $SetUserParams = @{
      Identity = $Username
      Company =  $OrganizationName
      Confirm = $false
    }

    # Sets delegate approval policy for office requests.
    if ($IsOffice)
    {
      $SetCalendarProcessingParams += @{
        AllRequestInPolicy = $true
        AllBookInPolicy = $false
        ForwardRequestsToDelegates = $true
        TentativePendingApproval = $true
        AddNewRequestsTentatively = $true
      }
    }

    ################################################################################################################  **Delegate and calendar permission settings for resource settings**  ##################################################################################################################
    
    # Authenticate as managed identity to Exchange Online if webhook trigger is used.
    if ($ValidWebhookData)
    {
      try {
        # Connect to Exchange Online using managed identity
        Connect-ExchangeOnline -ManagedIdentity -Organization $DomainName -ShowBanner:$false
      }
      catch {
        Write-Error -Message "$($Error[0].Exception.Message)"
        Exit 1
      }
    }

    # Authenticate as administrator to Exchange Online if parameters or JSON data is supplied.
    if ($ValidJson)
    {
      try {
        # Connect to Exchange Online as administrator.
        Connect-ExchangeOnline -Organization $DomainName -ShowBanner:$false
      }
      catch {
        Write-Error -Message "$($Error[0].Exception.Message)"
        Exit 1
      }
    }
    # Set parameters for Set-CalendarProcessing cmdlet based on if delegate is provided.
    if ($Delegates)
    {
      $ValidDelegates = $null
      $InvalidDelegates = $null

      foreach ($Delegate in $Delegates)
      {
        Write-Output "Looping through delegates for validation. Current delegate: $Delegate"
        try 
        {
            $OldPref = $global:ErrorActionPreference
            $global:ErrorActionPreference = 'Stop'

            # Validate delegate email address before setting processing rules.
            Get-EXOMailbox -Identity $Delegate | Out-Null

            # Concatenate valid delegate to string.
            $ValidDelegates = $ValidDelegates + $Delegate + ","
            
            # Loop through each permission provided in request
            foreach ($Permission in $CalendarPermissions)
            {
                # Set parameters for adding editor permissions to resource calendar.
                if ($Permission -eq "Editor")
                {
                    $EditorRights = $true
                    Write-Output "Editor permissions assigned to $Delegate"
                }

                # Add delegate approver to resource.
                if ($Permission -eq "Delegate")
                {
                    # Set delegate rights boolean to true.
                    $DelegateRights = $true
                    Write-Output "Delegate permissions assigned to $Delegate"
                }
            }
        }
        catch 
        {
            # Write error output to stream.
            Write-Error "Unable to find delegate $Delegate in Exchange. Skipping delegate assignment" # $Error[0].Exception.Message

            # Concatenate invalid delegate to string.
            $InvalidDelegates = $InvalidDelegates + $Delegate + ","
        }
        finally 
        {
            # Set global error action preference to default.
            $global:ErrorActionPreference = $OldPref
        }
      }

      # Check if delegate approver flag was added in request.
      if ($DelegateRights -and $ValidDelegates)
      {
          # Write to output stream results of check.
          Write-Output "Valid delegates: " + $ValidDelegates.Trim(',')

          # Parameters for Set-CalendarProcessing cmdlet.
          $SetCalendarProcessingParams += @{ ResourceDelegates = $ValidDelegates.Trim(',') }
      }
     }

    ################################################################################################################  **Runs cmdlets to set various settings defined in "Parameters for resource settings" and API calls**  #####################################################################################
    
    if ($ValidWebhookData)
    {
      # Sets API request body request based on success, failure, or warnings.
      $NewPrivateNoteSuccessBody = '{ "body":"<div>The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' has successfully created. <br><br> Please allow up to 24 hours for the resource to appear in Outlook Room Finder.</div>", "private":true }'
      $NewPrivateNoteFailureBody = '{ "body":"<div>The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' has failed to create. <br><br> Please reach out to your systems administrator for further assistance. Do <b>NOT</b> re-submit this request.</div>", "private":true }'
      $NewPrivateNoteResourceExistsBody = '{ "body":"The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' already exists. <br><br> Please check the information provided and try again by creating a new service request ticket.</div>", "private":true }'
      $InvalidDelegateBody = '{ "body":"<div>The delegate(s) ' + '<b>' + $InvalidDelegates.Trim(',') + '</b>' + ' do not contain valid email address(es).<br><br> Please reach out to your systems administrator for further assistance. Do <b>NOT</b> re-submit this request.</div>", "private":true }'
      $UpdateRequestedItemStatusCancelledBody = '{ "stage":3 }'
      $UpdateRequestedItemStatusFulfilledBody = '{ "stage":4 }'
    }
    # Check if identity exists before attempting operations. If no results are returned, proceed.
    if (!(Get-EXOMailbox -Identity $Username))
    {
    # Create resource mailbox.
    New-Mailbox @NewMailboxParams

    # Set Company Name attribute associated with resource.
    Set-User @SetUserParams

    # Set mailbox Display Name and type to Workspace.
    Set-Mailbox @SetMailboxParams

    # Set 30 second timer to allow resources to propagate prior to setting other resource values.
    Start-Sleep -Seconds 30

    # Set workspace details for location capacity, country, floor number, floor label, and wheelchair accessability.
    Set-Place @SetPlaceParams

    # Add Workspace as member to room list based on desginated Room List for an office.
    Add-DistributionGroupMember @AddDistributionGroupMemberParams

    # Set standard resource calendar processing rules.
    Set-CalendarProcessing @SetCalendarProcessingParams

    # Sets editor permissions on resource mailbox if flag for editor and approver rights are provided in the initial request.
    if ($ApproverRights -and $EditorRights)
    {

    }
    # Sets delegate permissions on resource mailbox if flag for approver rights are provided in the initial request.
    elseif ($ApproverRights)
    {

    }
    # Sets delegate permissions on resource mailbox if flag for approver rights are provided in the initial request.
    elseif ($EditorRights)
    {
      foreach ($Delegate in $ValidDelegates.Split(','))
      {
          try { Add-DistributionGroupMember -Identity $AdminGroup -Member $Delegate } catch { if ($Error[0].Exception.Message -match "Microsoft.Exchange.Management.Tasks.MemberAlreadyExistsException") { Write-Warning "User is already a member of group $AdminGroup" } else { Write-Error       Write-Error -Message "$($Error[0].Exception.Message)" } }
      }

      # Assigns appropriate mailbox permissions to admin group.
      Add-MailboxFolderPermission @AddMailboxParams -User $AdminGroup 
    }

################################################################################################################  **Post resource mailbox creation check**  #############################################################################################################################################

# Verify post mailbox creation.
if (Get-EXOMailbox -Identity $Username)
{
  if ($ValidWebhookData)
  {
  # Create private note with success status, and update requested item status to 'Fullfilled'.
    Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteSuccessBody -UseBasicParsing
    Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusFulfilledBody -UseBasicParsing
  }

  # Writes output to screen indicating that the resource creation was successful.
  Write-Output "Resource successfully created!"
}
else 
{
  if ($ValidWebhookData)
  {
    # Create private note with failure status and update requested item status to 'Cancelled'.
    Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteFailureBody -UseBasicParsing
    Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusCancelledBody -UseBasicParsing
  }
 }
}
else 
{
  if ($ValidWebhookData)
  {
    # Create private note indicating that resource already exists and update requested item status to 'Cancelled'.
    Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteResourceExistsBody -UseBasicParsing
    Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusCancelledBody -UseBasicParsing
  }

  # Write output to screen indicating that resource already exists
  Write-Output "Resource already exists. Please re-try using a unique indentifier to proceed."
}

if ($InvalidDelegates)
{
  if ($ValidWebhookData)
  {
    # Create private note indicating that the resource delegate was not applied due to an invalid email address.
    Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $InvalidDelegateBody -UseBasicParsing
  }

  # Write output to screen indicating that one or more resource delegates were not applied to the resource due to an invalid identity.
  Write-Output "Unable to assign one or more delegates due to an invalid identity."
}

# Disconnect from Exchange Online session.
Disconnect-ExchangeOnline -Confirm:$false