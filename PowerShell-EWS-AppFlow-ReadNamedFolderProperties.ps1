# PowerShell-EWS-AppFlow-ReadNamedFolderProperties.ps1


<#
.SYNOPSIS
  Uses EWS SOAP + OAuth client-credentials to return folder properties for the distinguished folder "recoverableitemsroot" in a target mailbox. 
  You can change the $NamedFolder variable to target other distinguished folders as needed.

.NOTES
  - Token scope for EWS app-only is typically: https://outlook.office365.com/.default
  - EWS endpoint: https://outlook.office365.com/EWS/Exchange.asmx
  - Keep X-AnchorMailbox aligned with the impersonated SMTP address.
  - Distinguished folder names: https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/ews-distinguished-folders
  - This sample returns the raw EWS response content. You can parse the XML to extract specific properties as needed.
  - For information on setting-up oAuth permissions for EWS in Azure AD, see:https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth
  - Please note that .com endpoints may need to be changed to .us, .de, .cn, etc. for sovereign clouds.
#>

# ==================================================
# 1) CONFIG for Authentication (fill these in)
# ==================================================
# You can specify the credentials directly in code here and comment out the injection line below for 
# PowerShell-EWS-AppFlow-ReadNamedFolderProperties_Credentials.ps1 or uncomment and set the three values below.
# This allows you to directly set the values here OR use the crentials from the seperate file. 
# Sometimes it helps to not show others your credentials while showing others your main code.
#$TenantId     = "YOUR_TENANT_ID"                # TODO - set if not using the injected credentials file
#$ClientId     = "YOUR_CLIENT_ID"                # TODO - set if not using the injected credentials file
#$ClientSecret = "YOUR_CLIENT_SECRET"            # TODO - set if not using the injected credentials file
. "$PSScriptRoot\PowerShell-EWS-AppFlow-ReadNamedFolderProperties_Credentials.ps1" 

# ============================================================
# 2) CONFIG for to point to the target mailbox + EWS endpoint
# ============================================================
$MailboxSmtpAddress = "user@contoso.com"   # TODO - set the target mailbox SMTP address
$EwsUrl = "https://outlook.office365.com/EWS/Exchange.asmx" # default EWS endpoint for O365 - change if needed for sovereign clouds or on-prem
$NamedFolder = "inbox" # TODO - set the distinguished folder name you want to target (e.g., "inbox", "calendar", "recoverableitemsroot", etc.)
#$NamedFolder = "recoverableitemsroot" # TODO - set the distinguished folder name you want to target (e.g., "inbox", "calendar", "recoverableitemsroot", etc.)
$displayCallAuditInfo = $true # Set to $true to display the raw EWS request and response content for auditing/troubleshooting purposes. Be cautious when enabling this in production environments due to potential sensitive information in the logs.
  # Set this to $false to display some resposne information adn the response body.
$UserAgent = "PowerShell-EWS-AppFlow-ReadNamedFolderProperties-Sample" # User-Agent header value to identify the application in EWS calls. This is optional but can be helpful for tracking and troubleshooting in EWS logs. You can set it to any string that makes sense for your application.

# ============================================================================
# 3) Get the token, call EWS, and display the raw response content
# ============================================================================

# ---- TLS (common in older PowerShell hosts) ----
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Get-EwsAppOnlyToken {
  param(
    [Parameter(Mandatory=$true)][string]$TenantId,
    [Parameter(Mandatory=$true)][string]$ClientId,
    [Parameter(Mandatory=$true)][string]$ClientSecret 
 
  )

  # Token endpoint (v2)
  $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

  # For EWS app-only, use the Outlook resource scope:
  # https://outlook.office365.com/.default
  $body = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
    scope         = "https://outlook.office365.com/.default"
  }

  $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -ContentType "application/x-www-form-urlencoded" -Body $body
  return $tokenResponse.access_token
}

 
function Invoke-EwsGetNamedFolderProperties {
  param(
    [Parameter(Mandatory=$true)][string]$EwsUrl,
    [Parameter(Mandatory=$true)][string]$AccessToken,
    [Parameter(Mandatory=$true)][string]$MailboxSmtpAddress,
    [Parameter(Mandatory=$true)][string]$NamedFolder,
    [Parameter(Mandatory=$false)][bool]$DisplayCallAuditInfo = $false
  )

  # SOAP for GetFolder(recoverableitemsroot)
  # BaseShape = AllProperties to return the full folder property bag EWS exposes.
  $soap = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016" />
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:SmtpAddress>$MailboxSmtpAddress</t:SmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
  </soap:Header>
  <soap:Body>
    <m:GetFolder>
             <m:FolderShape>
          <t:BaseShape>AllProperties</t:BaseShape>
          <t:AdditionalProperties>
            <t:FieldURI FieldURI="folder:DisplayName" />
          <t:ExtendedFieldURI PropertyTag="0x0E08" PropertyType="Integer" />   <!-- PR_MESSAGE_SIZE (0x0E08) -->
          </t:AdditionalProperties>
        </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="$NamedFolder">
          <t:Mailbox>
            <t:EmailAddress>$MailboxSmtpAddress</t:EmailAddress>
          </t:Mailbox>
        </t:DistinguishedFolderId>
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
"@

  # Headers:
  # - Authorization: Bearer <token>
  # - X-AnchorMailbox: MUST align with the mailbox being targeted (best practice for EXO routing).
  # - SOAPAction: GetFolder (some HTTP stacks require it)
  $ClientRequestId = [Guid]::NewGuid().ToString()  # Optional but good for tracking in logs
  $headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Accept"        = "text/xml"
    "X-AnchorMailbox" = $MailboxSmtpAddress
    "SOAPAction"    = "http://schemas.microsoft.com/exchange/services/2006/messages/GetFolder"
    "User-Agent"    = $UserAgent 
    "client-request-id" = $ClientRequestId  # Optional but good for tracking in logs
    "return-client-request-id" = "true"  # Echo back the client-request-id in the response for correlation
  }

  if ($DisplayCallAuditInfo) {
    # Display the raw EWS response content. You can parse the XML to extract specific properties as needed.
    Write-Host "Below is the request:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "POST: $EwsUrl" -ForegroundColor Cyan
 
    #$headers.GetEnumerator() | ForEach-Object { Write-Host ("{0}: {1}" -f $_.Key, $_.Value)  -ForegroundColor Cyan }
    $headers.GetEnumerator() | ForEach-Object { 
      #Write-Host ("{0}: {1}" -f $_.Key, $_.Value)  -ForegroundColor Cyan 
      if ($_.Key -eq "Authorization") {
        Write-Host "Authorization: <The Authorization header was redacted - it contains the Authentication information such as the Bearer token which is sensitive information.>" -ForegroundColor Red
      }
      else  
        {
          Write-Host ("{0}: {1}" -f $_.Key, $_.Value)  -ForegroundColor Cyan 
        }
      }
    Write-Host ""
    Write-Host $soap -ForegroundColor Cyan
    Write-Host ""
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    $StartUtc = Get-Date -AsUTC -Format "yyyy-MM-ddTHH:mm:ssZ"
    Write-Host "Making the EWS call...$($StartUtc)" -ForegroundColor Yellow
  }

  $resp = Invoke-WebRequest -Method Post -Uri $EwsUrl -Headers $headers -ContentType "text/xml; charset=utf-8" -Body $soap
  
  if ($DisplayCallAuditInfo) {
    $EndUtc = Get-Date -AsUTC -Format "yyyy-MM-ddTHH:mm:ssZ"
    Write-Host "EWS call completed at $($EndUtc)" -ForegroundColor Yellow  
    Write-Host "HTTP Status Code: $($resp.StatusCode)" -ForegroundColor Yellow 
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "EWS Response Headers:`n" -ForegroundColor Yellow
 
    $resp.Headers.GetEnumerator() | ForEach-Object { Write-Host ("{0}: {1}" -f $_.Key, $_.Value[0].ToString())  -ForegroundColor Green }

    Write-Host  ""
    $formattedResponse = Format-XML ([xml]$ewsResponse.Content)
    Write-Host $formattedResponse -ForegroundColor Green
  }

  return $resp
 
}

function Format-XML ([xml]$xml, $indent=2)
{
    $StringWriter = New-Object System.IO.StringWriter
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter
    $xmlWriter.Formatting = “indented”
    $xmlWriter.Indentation = $Indent
    $xml.WriteContentTo($XmlWriter)
    $XmlWriter.Flush()
    $StringWriter.Flush()
    return $StringWriter.ToString()
}
 
 
# ======================================
# ---------------- MAIN ----------------
# ======================================
cls
Write-Host "Running..." -ForegroundColor Green
Write-Host "----------------------------------------" -ForegroundColor Yellow
Write-Host "Getting the auth token..." -ForegroundColor Yellow
$token = Get-EwsAppOnlyToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
# Note: If you uncomment the line below then it will show the token. If you copy it into https://jwt.ms/ you can see the decoded token claims which is useful for troubleshooting.
#$token
# If Application permissions were granted then you should see "full_access_as_app", if you see "access_as_user" then it means Delegated permissions were granted 
# which won't work with this sample and you'll need to update the permissions in Azure AD and get admin consent for the app. Keep in mind that an Admin consent is required for Application permissions afer
# the permisisons were added via the manifest modifications in Azure.  
# For information on setting-up oAuth permissions for EWS in Azure AD, see:https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth

Write-Host "----------------------------------------" -ForegroundColor Yellow

Write-Host "Calling EWS to get the folder properties..." -ForegroundColor Yellow
$ewsResponse = Invoke-EwsGetNamedFolderProperties -EwsUrl $EwsUrl -AccessToken $token -MailboxSmtpAddress $MailboxSmtpAddress -NamedFolder $NamedFolder -DisplayCallAuditInfo $displayCallAuditInfo
 
# If the details of the call were not displayed above then display the raw EWS response content here. You can parse the XML to extract specific properties as needed.
if ($displayCallAuditInfo -eq $false) {
  Write-Host "----------------------------------------" -ForegroundColor Yellow
  Write-Host "Raw EWS response content:" -ForegroundColor Yellow  
  $ewsResponse  
  $ewsResponse.Content  
}
 
Write-Host "Finished"  -ForegroundColor Yellow
 
 
