<# 
.SYNOPSIS
Ingests webhook information and creates a reservable resource in Microsoft Exchange for a specified organization.

.DESCRIPTION 
A long description of how the script works and how to use it.

.NOTES 
Information about the environment, things to need to be consider and other information.

.COMPONENT 
Information about PowerShell Modules to be required.

.LINK 
Useful Link to ressources or others.

.Parameter ParameterName 
Description for a parameter in param definition section. Each parameter requires a separate description. The name in the description and the parameter section must match. 
#>

Param(
[parameter(Mandatory=$false)]
[Object]$WebhookData
)

#Import-Module ExchangeOnlineManagement

if ($WebhookData)
{  
# Outputs request header details.
Write-Output $WebhookData.RequestHeader

if ($WebhookData.RequestBody) 
{
# Converts request body to
$PayloadRequestBody = (ConvertFrom-Json -InputObject $WebhookData.RequestBody)

# Set PS variables for basic office attributes.
$Office = $PayloadRequestBody.location
$FloorNum = $PayloadRequestBody.floornum
$FloorLabel = $PayloadRequestBody.floorlabel
$Capacity = $PayloadRequestBody.capacity
$WheelChairAccessible = $PayloadRequestBody.wheelchairaccessible
$OfficeId = $PayloadRequestBody.officeid
$CubicleId = $PayloadRequestBody.cubicleid
$Delegates = $PayloadRequestBody.delegates.Split(',').Trim()
$CalendarPermissions = $PayloadRequestBody.calendarpermissions.Split(',')
$TicketID = $PayloadRequestBody.ticketid
$ServiceRequestItemID = $PayloadRequestBody.itemrequestid -replace '[\[\]]', ''

# Sets organization name, domain, and Azure subscription ID.
$OrganizationName = 'ORGANIZATION-NAME-HERE'
$FSDomain = 'FRESH-SERVICE-DOMAIN-HERE'
$DomainName = 'M365-DEFAULT-DOMAIN-HERE'
$SubscriptionId = 'AZURE-SUBSCRIPTION-ID-HERE'

# Sets Keyvault name, credential name, and administrative group variables for managing workspace resources.
$KeyvaultName = 'KEY-VAULT-NAME-HERE'
$CredentialName = 'CREDENTIAL-NAME-HERE'
$AdminGroup = 'MAIL-ENABLED-SECURITY-GROUP-HERE'

# Office 1 details
$Office1RoomList = 'roomlist1@contoso.com'
$Office1Building = 'Bldg 1'
$Office1Street = '1 Microsoft Way'
$Office1City = 'Redmond'
$Office1State = 'WA'
$Office1Zipcode = '13464'
$Office1Country = 'United States'

# Office 2 details
$Office2RoomList = 'roomlist2@contoso.com'
$Office2Building = 'Bldg 2'
$Office2Street = '1 Microsoft Way'
$Office2City = 'Redmond'
$Office2State = 'WA'
$Office2Zipcode = '13464'
$Office2Country = 'United States'

# Office 3 details
$Office3RoomList = 'roomlist2@contoso.com'
$Office3Building = 'Bldg 3'
$Office3Street = '1 Microsoft Way'
$Office3City = 'Redmond'
$Office3State = 'WA'
$Office3Zipcode = '13464'
$Office3Country = 'United States'

# Sets parameters for Get-AzKeyVaultSecret cmdlet to securely retrieve Freshservice Agent API creds for Freshservice API requests.
$KeyVaultParams = @{
Name = $CredentialName
VaultName = $KeyvaultName
AsPlainText = $true
}

# Set office and cubicle abbreviation values as well as resource type to be created.
$OfficeAbbr = 'OF'
$CubicleAbbr = 'WS'
$ResourceType = 'Workspace'

# Connect to Azure for retrieving credentials.
Connect-AzAccount -Subscription $SubscriptionId -Identity | Out-Null

# Sets API URLs while including unique ticket ID and service request item ID.
$FreshserviceCreatePrivateNoteUpdateURL = "https://$FSDomain/api/v2/tickets/$TicketID/notes"
$FreshserviceUpdateServiceRequestItemStatusURL = "https://$FSDomain/api/v2/tickets/$TicketID/requested_items/$ServiceRequestItemID"

# Sets header info for Freshservice API call. Retrieves Freshservice API key from Azure Key Vault and encodes using Base64.
$Headers = @{
  "Authorization" = ("Basic" + " " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('{0}:{1}' -f (Get-AzKeyVaultSecret @KeyVaultParams), $null))) )
  "Content-Type" = "application/json"
}

# Sets room list based on provided office name.
if ($Office -eq "OFFICE-NAME1")
{
  $RoomList = $Office1RoomList
  $Building = $Office1Building
  $Street = $Office1Street
  $City = $Office1City
  $State = $Office1State
  $Zipcode = $Office1Zipcode
  $Country = $Office1Country
}
elseif ($Office -eq "OFFICE-NAME2")
{
  $RoomList = $Office2RoomList
  $Building = $Office2Building
  $Street = $Office2Street
  $City = $Office2City
  $State = $Office2State
  $Zipcode = $Office2Zipcode
  $Country = $Office2Country
}
elseif ($Office -eq "OFFICE-NAME3")
{
  $RoomList = $Office3RoomList
  $Building = $Office3Building
  $Street = $Office3Street
  $City = $Office3City
  $State = $Office3State
  $Zipcode = $Office3Zipcode
  $Country = $Office3Country
}

# Set username using office abbreviation code and office ID #.
if ($OfficeId)
{
  $IsOffice = $true
  $Username = ($Building.ToLower() + "-" + $OfficeAbbr.ToLower() + "-" + $OfficeId)
  $DisplayName = ($OfficeAbbr + " " + $OfficeId.ToUpper())
}
if ($CubicleId)
{
  $Username = ($Building.ToLower() + "-" + $CubicleAbbr.ToLower() + "-" + $CubicleId)
  $DisplayName = ($CubicleAbbr + " " + $CubicleId.ToUpper())
}

################################################################################################################  **Parameters for resource settings**  ###################################################################################################################################################

# Base parameters for Set-Place cmdlet. Indicates basic resource location information.
$SetPlaceParams = @{
Identity = $Username
Building = $Building
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

# Connect to Exchange Online using managed identity
Connect-ExchangeOnline -ManagedIdentity -Organization $DomainName -ShowBanner:$false | Out-Null

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
          if ($Permission -eq "Editor (manage existing meetings)")
          {
              $EditorRights = $true
              Write-Output "Editor permissions assigned to $Delegate"
          }

          # Add delegate approver to resource.
          if ($Permission -eq "Approver")
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
}

################################################################################################################  **Runs cmdlets to set various settings defined in "Parameters for resource settings" and API calls**  #####################################################################################

# Sets API request body request based on success, failure, or warnings.
$NewPrivateNoteSuccessBody = '{ "body":"<div>The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' has successfully created. <br><br> Please allow up to 24 hours for the resource to appear in Outlook Room Finder.</div>", "private":true }'
$NewPrivateNoteFailureBody = '{ "body":"<div>The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' has failed to create. <br><br> Please reach out to your systems administrator for further assistance. Do <b>NOT</b> re-submit this request.</div>", "private":true }'
$NewPrivateNoteResourceExistsBody = '{ "body":"The resource ' + '<b>' + $DisplayName + ' (' + $Username + ')' + '</b>' + ' already exists. <br><br> Please check the information provided and try again by creating a new service request ticket.</div>", "private":true }'
$InvalidDelegateBody = '{ "body":"<div>The delegate(s) ' + '<b>' + $InvalidDelegates.Trim(',') + '</b>' + ' do not contain valid email address(es).<br><br> Please reach out to your systems administrator for further assistance. Do <b>NOT</b> re-submit this request.</div>", "private":true }'
$UpdateRequestedItemStatusCancelledBody = '{ "stage":3 }'
$UpdateRequestedItemStatusFulfilledBody = '{ "stage":4 }'

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
      try { Add-DistributionGroupMember -Identity $AdminGroup -Member $Delegate } catch { if ($Error[0].Exception.Message -match "Microsoft.Exchange.Management.Tasks.MemberAlreadyExistsException") { Write-Warning "User is already a member of group $AdminGroup" } else { Write-Error "Unable to add user $Delegate to admin group $AdminGroup" } }
  }

  # Assigns appropriate mailbox permissions to admin group.
  Add-MailboxFolderPermission @AddMailboxParams -User $AdminGroup 
}

################################################################################################################  **Post resource mailbox creation check**  #############################################################################################################################################

# Verify post mailbox creation.
if (Get-EXOMailbox -Identity $Username)
{
  # Create private note with success status, and update requested item status to 'Fullfilled'.
  Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteSuccessBody -UseBasicParsing
  Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusFulfilledBody -UseBasicParsing
}
else {
  # Create private note with failure status and update requested item status to 'Cancelled'.
  Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteFailureBody -UseBasicParsing
  Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusCancelledBody -UseBasicParsing
}
}
else {
# Create private note indicating that resource already exists and update requested item status to 'Cancelled'.
Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $NewPrivateNoteResourceExistsBody -UseBasicParsing
Invoke-WebRequest -Uri $FreshserviceUpdateServiceRequestItemStatusURL -Headers $Headers -Method Put -Body $UpdateRequestedItemStatusCancelledBody -UseBasicParsing
}

if ($InvalidDelegates)
{
# Create private note indicating that the resource delegate was not applied due to an invalid email address.
Invoke-WebRequest -Uri $FreshserviceCreatePrivateNoteUpdateURL -Headers $Headers -Method Post -Body $InvalidDelegateBody -UseBasicParsing
}

# Disconnect from Exchange Online session.
Disconnect-ExchangeOnline -Confirm:$false

#######################################################################################################################################################################################################################################################################################################
}
else {
Write-Output 'No webhook data received.'
}
