function Install-ZabbixAgent {
  <#
    .SYNOPSIS
        Install Zabbix agent for Windows
    .DESCRIPTION
        Downloads Zabbix agent and installs it for Windows x64 OS
    .EXAMPLE
        Install-ZabbixAgent
    .PARAMETER ZabbixServerIPAddress
        Your Zabbix Server IP Address, defaults to allow all connections
  #>
  [CmdletBinding()]
  param (
    [Parameter()]
    [string]$ZabbixServerIPAddress = '0.0.0.0/0'
  )

  # Download zabbix agent to here C:\Users\<user>\AppData\Local\Temp\2\ZabbixAgent.msi
  $outFile = Join-Path -Path $env:TEMP -ChildPath "ZabbixAgent.msi"
  $logFile = Join-Path -Path $env:TEMP -ChildPath "ZabbixAgentInstallLog.txt"
  $agentFileExist = Test-Path $outFile
  $zabbixAgentInstalled = Get-Service 'Zabbix Agent' -ErrorAction SilentlyContinue

  if (!$agentFileExist) {
    $ProgressPreference = 'SilentlyContinue'
    Write-Verbose 'Downloading Zabbix agent 5.0.3 from www.zabbix.com'
    Invoke-WebRequest -Uri "https://www.zabbix.com/downloads/5.0.3/zabbix_agent-5.0.3-windows-amd64-openssl.msi" -OutFile $outFile
  }

  if (!$zabbixAgentInstalled) {
    Write-Verbose 'Installing Zabbix Agent'
    Start-Process -FilePath $outFile -ArgumentList "/l*v $logFile SERVER=$ZabbixServerIPAddress /qb!" -PassThru -Wait
  }

  $agentIsRunning = Test-NetConnection localhost -Port 10050 -InformationLevel Quiet
  Write-Verbose "Zabbix agent is downloaded and installed to $($env:computername). Agent is running on port 10050: $agentIsRunning"
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
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]$ZabbixHost,
    [Parameter(Mandatory)]
    [string]$Token,
    [Parameter(Mandatory)]
    [string]$AgentIPAddress
  )

  $ctype = "application/json"
  $url = "http://$($ZabbixHost)/api_jsonrpc.php"

  # 10081 = Template OS Windows by Zabbix agent
  $templateId = 10081

  # 10 = Templates/Operating systems, 6 = Virtual machines
  $groupId = 10

  # 10050 = default
  $zabbixPort = "10050"

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
          dns   = ""
          port  = $zabbixPort
        }
      )
      groups     = @(
        @{ groupid = $groupId }
      )
      templates  = @(
        @{ templateid = $templateId }
      )
    }
  } | ConvertTo-Json -Depth 3

  $res = Invoke-RestMethod -Uri $url -Body $createHostBody -Method Post -ContentType $ctype
  $res.result.hostids
}
