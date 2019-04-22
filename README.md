# Posh-Notify

Send yourself or your team notifications from within a powershell script. Alert when a task is done or something needs attention. Keep a list of all actions.

## Usage

```powershell
# get the module
Install-Module Posh-Notify

# set a default channel to send to
$Global:TeamsWebhook = "<webhook from teams channel>"

# use
Notify-Teams "test"
Get-ChildItem | Notify-Teams -Webhook "<different channel's webhook>"
# etc
```
