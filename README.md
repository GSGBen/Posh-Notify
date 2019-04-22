# Posh-Notify

Send yourself or your team notifications from within a powershell script. Alert when a task is done or something needs attention. Keep a list of all actions. Play audible beeps when a local task has finished.

## Usage

```powershell
# get the module
Install-Module Posh-Notify

# set a default channel to send to
$Global:TeamsWebhook = "<webhook from teams channel>"
# set a default mailbox to send to
$Global:OutlookWebhook = "<webhook from OWA (see below)>"

# send messages
Notify-Teams "test"
Get-ChildItem | Notify-Teams -Webhook "<different channel's webhook>"
Notify-Outlook "test"
# etc

# play sounds on completion
Start-LongTask; Beep 2000 500
Wait-ForThingsToFinish; Arpeggio
```

## Getting a Teams webhook

Right click on the channel -> connectors -> add webhook and copy URL

## Getting an Outlook Webhook

Office 365 only. Log into https://outlook.office365.com/owa -> gear icon in the top right -> manage connectors > add webhook and copy URL.
