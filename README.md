# Zabbix host Powershell Module

![Publish](https://github.com/Satak/ps-zabbix-host/workflows/Publish/badge.svg)
[![PS Gallery][psgallery-badge-dt]][powershell-gallery]
[![PS Gallery][psgallery-badge-v]][powershell-gallery]

![zabbix logo](https://raw.githubusercontent.com/Satak/ps-zabbix-host/master/icon/zabbix-icon-48x48.png 'Zabbix logo')

Powershell module `PS-Zabbix-Host` to install and register Zabbix agent/host on Windows

| Version | Info            | Date       |
| ------- | --------------- | ---------- |
| 0.0.4   | Initial release | 11.09.2020 |

## Commands

| Command               | Info                                                      |
| --------------------- | --------------------------------------------------------- |
| `Install-ZabbixAgent` | Downloads Zabbix agent and installs it for Windows x64 OS |
| `New-ZabbixToken`     | Login to Zabbix server and get token                      |
| `New-ZabbixHost`      | Use Zabbix REST API to create a new host (agent)          |
| `Get-LocalIPAddress`  | Get IPv4 IP Address for creating new a Zabbix host        |

## Usage

```powershell
param(
    $username = 'Admin',
    $password = 'zabbix',
    $zabbixHost = '10.0.0.2'
)

Install-Module -Name PS-Zabbix-Host -Force -Confirm:$False

# Zabbix server credentials
$credentials = New-Object System.Management.Automation.PSCredential(
  $userName,
  (ConvertTo-SecureString $password -AsPlainText -Force)
)

# download agent msi package from www.zabbix.com and install it
Install-ZabbixAgent

# IP address of the client where the agent is running
$ip = Get-LocalIPAddress

# Zabbix server token
$token = New-ZabbixToken -ZabbixHost $zabbixHost -Credentials $credentials

# create new host by using the REST API and token
New-ZabbixHost -ZabbixHost $zabbixHost -Token $token -AgentIPAddress $ip
```

[powershell-gallery]: https://www.powershellgallery.com/packages/PS-Zabbix-Host/
[psgallery-badge-dt]: https://img.shields.io/powershellgallery/dt/PS-Zabbix-Host.svg
[psgallery-badge-v]: https://img.shields.io/powershellgallery/v/PS-Zabbix-Host.svg
