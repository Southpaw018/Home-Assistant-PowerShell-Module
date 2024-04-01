<#
.SYNOPSIS
  Name: New-HomeAssistantSession.ps1
 This is a function in the a PowerShell Module to Control the Home-Assistant home automation software.
.DESCRIPTION
  This function is used to connect to the Home-Assistent API
.NOTES
    Updated in 2020
  Author:Bevin Du Plessis
.LINK
  https://github.com/bevinduplessis
.EXAMPLE
 New-HomeAssistantSession -ip <Home-Assistant IP> -port <Home-Assistant port> -token <access token>
#>

Function New-HomeAssistantSession {
  Param(
    [Parameter(Mandatory = $true)][string]$Hostname,
    [Parameter(Mandatory = $true)][string]$Port,
    [Parameter(Mandatory = $true)][string]$Token,
    [Switch]$UseSSL
  )

  $global:ha_api_headers = @{Authorization = "Bearer " + $Token }
    
  if ($UseSSL.IsPresent) {
    $global:ha_api_url = "https://" + "$Hostname" + ":" + "$Port" + "/api/"
  }
  else {
    $global:ha_api_url = "http://" + "$Hostname" + ":" + "$Port" + "/api/"
  }

  try {
    Write-Host "Testing connection... $global:ha_api_url " -ForegroundColor Yellow -NoNewline
    $api_connection = (Invoke-WebRequest -uri $ha_api_url -Method GET -Headers $ha_api_headers -UseBasicParsing)    
    $global:ha_api_configured = $true
    Write-Host "Connection to Home-Assistant API succeeded! (" ($api_connection).StatusCode ($api_connection).StatusDescription ")" -ForegroundColor Green
  }
  catch {

    if ((Test-NetConnection -ComputerName $Hostname -WarningAction SilentlyContinue).PingSucceeded) {
      Write-Host "Connection to Home-Assistant API failed!" -ForegroundColor Red
    }
    else {
      Write-Host "Connection failed - ICMP request timed out!" -ForegroundColor Red
    }
    
    $global:ha_api_url = $null
    $global:ha_api_headers = $null
    $global:ha_api_configured = $false
  } #catch end

} #End of function New-HomeAssistantSessiontion New-HomeAssistantSession

<#
.SYNOPSIS
  Name: Invoke-HomeAssistantService.ps1
 This is a function in the a PowerShell Module to Control the Home-Assistant home automation software.
.DESCRIPTION
  This function is used to invoke a service in the Home-Assistent
.NOTES
    Original release Date: 17.03.2019
  Author: Flemming Sørvollen Skaret (https://github.com/flemmingss/)
.LINK
  https://github.com/flemmingss/
.EXAMPLE
 Invoke-HomeAssistantService -service <service> -entity_id <entity_id>
 Invoke-HomeAssistantService -service <service> -json '<json>'
#>

Function Invoke-HomeAssistantService {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true, Position = "0", ParameterSetName = "ByEntity")][string]
    [Parameter(Mandatory = $true, Position = "0", ParameterSetName = "ByJSON")][string]
    $service,

    [Parameter(Mandatory = $true, Position = "1", ParameterSetName = "ByEntity")][string]
    $entity_id,

    [Parameter(Mandatory = $true, Position = "1", ParameterSetName = "ByJSON")][string]
    $json
  )

  #-SERVICE
  $service = $service.replace(".", "/")

  #-ENTITY_ID
    
  if ($entity_id) {
    $body = ConvertTo-Json @{"entity_id" = "$entity_id" }
  }

  #-JSON
  elseif ($json) {
    $body = $json    
  }

  Invoke-RestMethod -Method Post -Uri ("$ha_api_url" + "services/" + "$service") -Body ($body) -Header $ha_api_headers
  Write-Host $body
} #End of function Invoke-HomeAssistantService

<#
.SYNOPSIS
  Name: Get-HomeAssistantConfig.ps1
 This is a function in the a PowerShell Module to Control the Home-Assistant home automation software.
.DESCRIPTION
  This function is used to get information abaout the Home-Assistant configuration
.NOTES
    Original release Date: 17.03.2019
  Author: Flemming Sørvollen Skaret (https://github.com/flemmingss/)
.LINK
  https://github.com/flemmingss/
.EXAMPLE
  Get-HomeAssistantConfig
#>

function Get-HomeAssistantConfig {
  Invoke-RestMethod -Method get -Uri ("$ha_api_url" + "config") -Header $ha_api_headers
} #End of function Get-HomeAssistantConfig

<#
.SYNOPSIS
  Name: Get-HomeAssistantEntity.ps1
 This is a function in the a PowerShell Module to Control the Home-Assistant home automation software.
.DESCRIPTION
  This function is used to get information abaout a entity in Home-Assistant
.NOTES
    Original release Date: 17.03.2019
  Author: Flemming Sørvollen Skaret (https://github.com/flemmingss/)
.LINK
  https://github.com/flemmingss/
.EXAMPLE
  Get-HomeAssistantEntity
  Get-HomeAssistantEntity -entity_id <entity_id>
#>

function Get-HomeAssistantEntity {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $false, Position = "0")
    ][string]$entity_id
  )

  if ($entity_id) {
    Invoke-RestMethod -Method get -Uri ("$ha_api_url" + "states/" + $entity_id) -Header $ha_api_headers
  }
  else {
    Invoke-RestMethod -Method get -Uri ("$ha_api_url" + "states") -Header $ha_api_headers
  }

} #End of function Get-HomeAssistantEntity

<#
.SYNOPSIS
  Name: Get-HomeAssistantServices.ps1
 This is a function in the a PowerShell Module to Control the Home-Assistant home automation software.
.DESCRIPTION
  This function is used to get all the available services in Home-Assistant
.NOTES
    Original release Date: 17.03.2019
  Author: Flemming Sørvollen Skaret (https://github.com/flemmingss/)
.LINK
  https://github.com/flemmingss/
.EXAMPLE
  Get-HomeAssistantServices
#>

Function Get-HomeAssistantServices { 
  Invoke-RestMethod -Method get -Uri ("$ha_api_url" + "services") -Header $ha_api_headers
} #End of function Get-HomeAssistantServices

Function Get-HomeAssistantAreas {

  Invoke-RestMethod -Method Post -Uri ("$ha_api_url" + "template") -Body ('{"template": "{{ areas() }}"}') -Header $ha_api_headers
} #

function Get-HomeAssistantEntityByArea {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $false, Position = "0")]      
    [string]$Area,
    [Parameter(Mandatory = $false, Position = "1")]
    [AllowEmptyString()]
    [ValidateSet('binary_sensor', 'button', 'climate', 'cover', 'device_tracker', 'fan', 'image', 'light', 'lock', 'media_player', 'number', 'remote', 'scene', 'select', 'sensor', 'switch', 'time', 'update', 'vacuum', 'weather', '')]
    [string]$Type 
  
  )

  $Area = $Area.ToLower()
  $Type = $Type.ToLower()

  if ($Area) {      
      
    $template = '{"template": "{{ area_entities(\"_AREA_\") | select(\"match\", \"_TYPE_\") | list }}"}' -replace '_AREA_', "$Area" -replace '_TYPE_', "$Type"
    Invoke-RestMethod -Method Post -Uri ("$ha_api_url" + "template") -Body ($template) -Header $ha_api_headers 
  }
  else {        
    $Areas = Get-HomeAssistantAreas 
    $Areas | ForEach-Object {
      if ($Type) {
        "--> $_"
        $template = '{"template": "{{ area_entities(\"_AREA_\") | select(\"match\", \"_TYPE_\") | list }}"}' -replace '_AREA_', "$_" -replace '_TYPE_', "$Type"
      }
      Else {
        "--> $_"
        $template = '{"template": "{{ area_entities(\"_AREA_\")  | list }}"}' -replace '_AREA_', "$_"
      }
      Invoke-RestMethod -Method Post -Uri ("$ha_api_url" + "template") -Body ($template) -Header $ha_api_headers
          
          
    }
      
  }

} #End of function Get-HomeAs