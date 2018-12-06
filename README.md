# Posh-Notify

Send yourself or your team notifications from within a powershell script. Alert when a task is done or something needs attention. Keep a list of all actions.

## Usage

```
Install-Module Posh-Notify

$TeamsWebhook = <webhook from teams channel>

Notify-Teams "test"
Get-ChildItem | Notify-Teams
etc
```