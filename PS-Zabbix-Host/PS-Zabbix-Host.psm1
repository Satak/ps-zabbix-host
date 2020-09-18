function Install-ZabbixAgent {
  <#
    .SYNOPSIS
        Install Zabbix agent for Windows
    .DESCRIPTION
        Downloads Zabbix agent and installs it for Windows x64 OS
    .EXAMPLE
        Install-ZabbixAgent
    .PARAMETER ZabbixServer
        List of comma delimited IP addresses (or hostnames) of Zabbix servers and Zabbix proxies. Spaces are allowed. Defaults to accept all.
    .PARAMETER ZabbixServerActive
        IP:port (or hostname:port) of Zabbix server or Zabbix proxy for active checks. If port is not specified, default port 10051 is used.
    .PARAMETER ListenPort
        Agent will listen on this port for connections from the server. Default port: 10050
    .PARAMETER AgentUrl
        URL where to download Zabbix Agent. Default url: https://www.zabbix.com/downloads/5.0.3/zabbix_agent-5.0.3-windows-amd64-openssl.msi
    .PARAMETER LocalAgentInstaller
        Set full path to local agent installer e.g. C:\Windows\Temp\zabbixAgent.msi
  #>
  [CmdletBinding()]
  param (
    [Parameter()]
    [string]$ZabbixServer = '0.0.0.0/0',
    [Parameter()]
    [string]$ZabbixServerActive,
    [Parameter()]
    [string]$ListenPort = '10050',
    [Parameter()]
    [string]$AgentUrl = 'https://www.zabbix.com/downloads/5.0.3/zabbix_agent-5.0.3-windows-amd64-openssl.msi',
    [Parameter()]
    [string]$LocalAgentInstaller
  )

  if ($LocalAgentInstaller) {
    $outFile = $LocalAgentInstaller
    Write-Verbose "Using local installer from $outFile"
  }
  else {
    # set to download zabbix agent to: C:\Users\<user>\AppData\Local\Temp\2\ZabbixAgent.msi
    $outFile = Join-Path -Path $env:TEMP -ChildPath "ZabbixAgent.msi"
    Write-Verbose "Download path set to $outFile"
  }

  $logFile = Join-Path -Path $env:TEMP -ChildPath "ZabbixAgentInstallLog.txt"
  $agentFileExist = Test-Path $outFile
  $zabbixAgentInstalled = Get-Service 'Zabbix Agent' -ErrorAction SilentlyContinue

  Write-Verbose "Pre-install checks: Zabbix agent installer exists: $agentFileExist Agent is already installed: $($zabbixAgentInstalled.name)"
  $ProgressPreference = 'SilentlyContinue'

  # download zabbix agent if it's not already locally available
  if (!$agentFileExist) {
    Write-Verbose "Downloading Zabbix agent from $AgentUrl"
    Invoke-WebRequest -Uri $AgentUrl -OutFile $outFile
  }

  # install Zabbix agent if it's not installed
  if (!$zabbixAgentInstalled) {
    $zabbixOptions = "SERVER=$($ZabbixServer.Replace(' ','')) ListenPort=$ListenPort"
    if ($ZabbixServerActive) {
      $zabbixOptions += " SERVERACTIVE=$($ZabbixServerActive.Replace(' ',''))"
    }
    $zabbixAgentArgs = "/l*v $logFile $zabbixOptions /qb!"
    Write-Verbose "Installing Zabbix Agent with args $zabbixAgentArgs"
    Start-Process -FilePath $outFile -ArgumentList $zabbixAgentArgs -PassThru -Wait
  }

  $agentIsRunning = Test-NetConnection localhost -Port $ListenPort -InformationLevel Quiet
  Write-Verbose "Zabbix agent is downloaded and installed to $($env:computername). Agent is running on port $($ListenPort): $agentIsRunning"
}

function Get-LocalIPAddress {
  <#
    .SYNOPSIS
        Get client's local IP address
    .DESCRIPTION
        Get IPv4 IP Address for creating new a Zabbix host
    .EXAMPLE
        Get-LocalIPAddress
  #>
  (
    Get-NetIPConfiguration |
    Where-Object {
      $null -ne $_.IPv4DefaultGateway -and
      $_.NetAdapter.Status -ne "Disconnected"
    }
  ).IPv4Address.IPAddress
}

function New-ZabbixToken {
  <#
    .SYNOPSIS
        Get auth token from Zabbix server
    .DESCRIPTION
        Login to Zabbix server and get token
    .EXAMPLE
         New-ZabbixToken -ZabbixHost "my-zabbix-instance" -Credentials (Get-Credential)
    .PARAMETER ZabbixHost
        Your Zabbix Server IP Address/DNS name. URL is constructed from this "http://$($ZabbixHost)/api_jsonrpc.php"
    .PARAMETER Credentials
        PSCredential for authenticating to Zabbix server REST API
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$ZabbixHost,
    [Parameter(Mandatory)]
    [PSCredential]$Credentials
  )

  $ctype = "application/json"
  $url = "http://$($ZabbixHost)/api_jsonrpc.php"

  $loginBody = @{
    jsonrpc = "2.0"
    method  = "user.login"
    id      = 1
    auth    = $null
    params  = @{
      user     = $Credentials.UserName
      password = $Credentials.GetNetworkCredential().Password
    }
  } | ConvertTo-Json

  $login = Invoke-RestMethod -Uri $url -Body $loginBody -Method Post -ContentType $ctype
  $token = $login.result
  $token
}

function New-ZabbixHost {
  <#
    .SYNOPSIS
        Create new Zabbix host
    .DESCRIPTION
        Use Zabbix REST API to create a new host (agent)
    .EXAMPLE
        New-ZabbixHost -ZabbixHost "my-zabbix-instance" -Token "bb9095c09d2037ac65126d13965d9fb3" -AgentIPAddress "10.0.0.3"
    .PARAMETER ZabbixHost
        Your Zabbix Server IP Address/DNS name. URL is constructed from this "http://$($ZabbixHost)/api_jsonrpc.php"
    .PARAMETER Token
        Token from New-ZabbixToken cmdlet
    .PARAMETER AgentIPAddress
        Client/Host local IP address that is reachable from Zabbix server (use Get-LocalIPAddress cmdlet)
    .PARAMETER TemplateId
        Zabbix template Ids to assign to the host. Accepts one or more values.
        Template OS Windows by Zabbix agent = 10081 (default)
        Template OS Windows by Zabbix agent active = 10299
        Template OS Linux by Zabbix agent = 10001
        Template OS Linux by Zabbix agent active = 10284
    .PARAMETER GroupId
        Zabbix group Ids to assign to the host. Accepts one or more values. Default: 10 = Templates/Operating systems. Alternative: 6 = Virtual machines
    .PARAMETER AgentPort
        Zabbix agent port. Default 10050
    .PARAMETER DNSName
        Zabbix agent dns name
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$ZabbixHost,
    [Parameter(Mandatory)]
    [string]$Token,
    [Parameter(Mandatory)]
    [string]$AgentIPAddress,
    [string[]]$TemplateId = 10081,
    [string[]]$GroupId = 10,
    [string]$AgentPort = "10050",
    [string]$DNSName = ""
  )

  $ctype = "application/json"
  $url = "http://$($ZabbixHost)/api_jsonrpc.php"
  [array]$groups = foreach ($group in $GroupId) { @{groupid = $group } }
  [array]$templates = foreach ($template in $TemplateId) { @{templateid = $template } }

  $createHostBody = @{
    jsonrpc = "2.0"
    method  = "host.create"
    id      = 1
    auth    = $Token
    params  = @{
      host       = $env:computername
      interfaces = @(
        @{type  = 1
          main  = 1
          useip = 1
          ip    = $AgentIPAddress
          dns   = $DNSName
          port  = $AgentPort
        }
      )
      groups     = $groups
      templates  = $templates
    }
  } | ConvertTo-Json -Depth 3

  $res = Invoke-RestMethod -Uri $url -Body $createHostBody -Method Post -ContentType $ctype
  $res.result.hostids
  Write-Verbose "Zabbix host ($AgentIPAddress) registered to server $ZabbixHost payload: $createHostBody"
}
