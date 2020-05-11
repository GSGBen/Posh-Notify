<#
.synopsis
    * Posts a message to a Microsoft Office 365 service via webhook
.description
    * Posts a message to a Microsoft Office 365 service via webhook
    * Not sure of the max character limit, so truncating is disabled
    * Supports Text and Title via argument. Other Office 365 card attributes can be added via properties in a hash table passed through to the AdditionalProperties argument
      * See https://messagecardplayground.azurewebsites.net/ for the valid attributes, with examples
    * No default webhook - these are provided by wrapper functions like Send-TeamsNotification and Send-OutlookNotification
.parameter Text
    * The text to display
    * Array, so works for array or single string
      * i.e. will send a table as easily as a line of text
    * Will concatenate array as new lines, not spaces
.parameter Webhook
    * The Webhook URL of the service to post to
.parameter AdditionalProperties
    * Construct a fancier card via additional options
      * See https://messagecardplayground.azurewebsites.net/ for the valid attributes, with examples
.notes
    * TODO:
      * Support \r\n line endings
      * Implement AdditionalProperties
    * Author: Ben Renninson
    * Email: ben@goldensyrupgames.com
    * From: https://github.com/GSGBen/Posh-Notify
#>
function Send-Office365Notification
{
    Param
    (
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)][string[]]$Text,
        [Parameter(Position=1,Mandatory=$false)][string]$Title,
        [Parameter(Position=2,Mandatory=$false)][string]$Webhook,
        [Parameter(Position=3,Mandatory=$false)][System.Collections.IDictionary]$AdditionalProperties
    )

    #Join each element that may be passed through in the pipeline
    Process{
        #If passed through the pipeline, join to a single array
        # after making sure there are no empty lines,
        # and breaking up multi-line strings - the formatting below needs an array of single strings
        $CombinedArray += $Text -replace '\n\n',"`n `n" -split "`n"
    }
    #only send after joining
    End{ 

        #Specified webhook or fail
        if ($Webhook)
        {
            $FinalWebhook = $Webhook
        }
        else
        {
            Write-Error "No webhook specified and `$TeamsWebhook not set!"
            return
        }

        #Join combined array into multi-line text. Need to restart the code formatting wraps each time. Also need two spaces in front of the newline character (which should show up as \n to teams, via the json conversion).
        $JoinString = '```  ' + "`n" + ' ```'
        $FinalText += $CombinedArray -join $JoinString
    
        #construct body as hash table.
        #Add starting and ending code formatting wraps. Internal ones are above with the array combination 
        $Body = @{
            "title" = $Title
            "text" = '```' + $FinalText + '```'
        }

        #Add the addition properties if they were used
        if ($AdditionalProperties)
        {
            $Body = Merge-HashTable $Body $AdditionalProperties
        }

        #Convert hashtable to json
        $BodyAsJson = ConvertTo-Json $Body

        #Send the text to the webhook
        Invoke-RestMethod -ContentType "application/json" -Body $BodyAsJson -Method Post -URI $FinalWebhook
    }
}

<#
.synopsis
    * Posts a message to a Microsoft Teams channel via webhook
.description
    * Posts a message to a Microsoft Teams channel via webhook
    * Not sure of the max character limit, so truncating is disabled
    * Supports Text and Title via argument. Other Office 365 card attributes can be added via properties in a hash table passed through to the AdditionalProperties argument
      * See https://messagecardplayground.azurewebsites.net/ for the valid attributes, with examples
    * Selects the webhook to use in the following priority (first most preferred) order, with the aim of easiest use
      * Webhook argument passed to function
      * $TeamsWebhook
        * e.g. set $Global:TeamsWebhook in your powershell profile as a default
      * if none of the above, it errors. No prompting because of the automation focus
.parameter Text
    * The text to display
    * Array, so works for array or single string
      * i.e. will send a table as easily as a line of text
    * Will concatenate array as new lines, not spaces
.parameter Webhook
    * The Webhook URL of the channel to post to
    * To get this
      * right click on the channel and select 'connectors'
      * edit an existing webhook connector and get the url or
      * add a new 'webhook' connector them get the url
.parameter AdditionalProperties
    * Construct a fancier card via additional options
      * See https://messagecardplayground.azurewebsites.net/ for the valid attributes, with examples
.notes
    * TODO:
      * Implement AdditionalProperties
    * Author: Ben Renninson
    * Email: ben@goldensyrupgames.com
    * From: https://github.com/GSGBen/Posh-Notify
#>
function Send-TeamsNotification
{
    Param
    (
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)][string[]]$Text,
        [Parameter(Position=1,Mandatory=$false)][string]$Title,
        [Parameter(Position=2,Mandatory=$false)][string]$Webhook,
        [Parameter(Position=3,Mandatory=$false)][System.Collections.IDictionary]$AdditionalProperties
    )

    #Specified webhook or global or fail
    if ($Webhook)
    {
        $FinalWebhook = $Webhook
    }
    elseif ($TeamsWebhook)
    {
        $FinalWebhook = $TeamsWebhook
    }
    else
    {
        Write-Error "No webhook specified and `$TeamsWebhook not set!"
        return
    }

    Send-Office365Notification -Text $Text -Title $Title -Webhook $FinalWebhook -AdditionalProperties $AdditionalProperties
}

<#
.synopsis
    * Posts a message to a Microsoft Office 365 Outlook mailbox via webhook
.description
    * Posts a message to a Microsoft Office 365 Outlook mailbox via webhook
    * Not sure of the max character limit, so truncating is disabled
    * Supports Text and Title via argument. Other Office 365 card attributes can be added via properties in a hash table passed through to the AdditionalProperties argument
      * See https://messagecardplayground.azurewebsites.net/ for the valid attributes, with examples
    * Selects the webhook to use in the following priority (first most preferred) order, with the aim of easiest use
      * Webhook argument passed to function
      * $OutlookWebhook
        * e.g. set $Global:OutlookWebhook in your powershell profile as a default
      * if none of the above, it errors. No prompting because of the automation focus
.parameter Text
    * The text to display
    * Array, so works for array or single string
      * i.e. will send a table as easily as a line of text
    * Will concatenate array as new lines, not spaces
.parameter Webhook
    * The Webhook URL of the mailbox to post to
    * To get this
      * log into https://outlook.office365.com/owa
      * Select the gear icon in the top right hand corner
      * manage connectors
      * add a new 'webhook' connector them get the url
.parameter AdditionalProperties
    * Construct a fancier card via additional options
      * See https://messagecardplayground.azurewebsites.net/ for the valid attributes, with examples
.notes
    * TODO:
      * Implement AdditionalProperties
    * Author: Ben Renninson
    * Email: ben@goldensyrupgames.com
    * From: https://github.com/GSGBen/Posh-Notify
#>
function Send-OutlookNotification
{
    Param
    (
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)][string[]]$Text,
        [Parameter(Position=1,Mandatory=$false)][string]$Title,
        [Parameter(Position=2,Mandatory=$false)][string]$Webhook,
        [Parameter(Position=3,Mandatory=$false)][System.Collections.IDictionary]$AdditionalProperties
    )

    #Specified webhook or global or fail
    if ($Webhook)
    {
        $FinalWebhook = $Webhook
    }
    elseif ($TeamsWebhook)
    {
        $FinalWebhook = $OutlookWebhook
    }
    else
    {
        Write-Error "No webhook specified and `$OutlookWebhook not set!"
        return
    }

    Send-Office365Notification -Text $Text -Title $Title -Webhook $FinalWebhook -AdditionalProperties $AdditionalProperties
}

<#
.synopsis
    - Posts a message to a Discord channel via webhook
.description
    - Posts a message to a Discord channel via webhook
    - Discord webhooks specify both the user posting it and the channel
    - Max is 2000 chars
    - By default this function will send the first 2000 characters of the Text input (minus characters required for formatting, prefix and suffix). If you specify -LastChars it'll send the last 2000
    - Should support Discord markdown formatting
    - Uses $Global:DiscordDefaultWebhook if no Webhook is specified
.parameter Text
    - The text to display
    - Array, so works for array or single string
    - Will concatenate array as new lines, not spaces
.parameter Prefix
    - Comes before Text. Not truncated at the moment, so don't make this + Suffix anywhere close to 2000 
.parameter Suffix
    - Comes after Text. Not truncated at the moment, so don't make this + Prefix anywhere close to 2000 
.parameter Webhook
    - The Webhook URL of the channel to post to
.parameter LastChars
    - By default this function will send the first 2000 characters (minus characters required for formatting). If you specify -LastChars it'll send the last 2000
    - E.g. if you want as much info as possible from some output but definitely want the final result
.parameter NoSeparators
    - By default we separate messages with some basic formatting (e.g. ***). Specify this to not
.notes
    Author: Ben Renninson
    Email: ben@goldensyrupgames.com
    From: https://github.com/GSGBen/Posh-Notify
#>
function Send-DiscordNotification
{
    Param
    (
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)][string[]]$Text,
        [Parameter(Position=1,Mandatory=$false)][string[]]$Prefix,
        [Parameter(Position=2,Mandatory=$false)][string[]]$Suffix,
        [Parameter(Position=3,Mandatory=$false)][string]$Webhook,
        [Parameter(Position=4,Mandatory=$false)][switch]$LastChars = $false,
        [Parameter(Position=5,Mandatory=$false)][switch]$NoSeparators = $false
    )

    #Specified webhook or global or fail
    if ($Webhook)
    {
        $FinalWebhook = $Webhook
    }
    elseif ($Global:DiscordDefaultWebhook)
    {
        $FinalWebhook = $Global:DiscordDefaultWebhook
    }
    else
    {
        Write-Error "No webhook specified and `$Global:DiscordDefaultWebhook not set!"
        return
    }

    #Separate messages because if there are multiple from the same user quickly, it looks like one
    $StartSeparator = "**---**`n"
    $EndSeparator = "`n**---**"

    # Convert the text, which may be an array of strings, to a single string. Join with new lines
    $FinalText = $Text -join "`r`n"
    $FinalPrefix = $Prefix -join "`r`n"
    $FinalSuffix = $Suffix -join "`r`n"

    #region TRUNCATION--------------------------------------------------------

        # Specified by discord
        $MaxLength = 2000
        # Adjust Text input to fit. Truncate to 0 if required, but no lower
        $UsableLength = [math]::max($MaxLength - $StartSeparator.Length - $EndSeparator.Length - $FinalPrefix.Length - $FinalSuffix.Length,0)

        # Cut the text if it's greater than 2000, taking into account all other formatting
        if ($FinalText.Length -gt 2000)
        {
            if ($LastChars) #select from end
            {
                $FinalText = $FinalText.Substring(($FinalText.Length-$UsableLength),$UsableLength)
            }
            else # select from start
            {
                $FinalText = $FinalText.Substring(0,$UsableLength)
            }
        }
        else
        {
            # It's within the allowed character limit. Do nothing
        }

    #endregion TRUNCATION--------------------------------------------------------

    # add prefix and suffix
    $FinalText = $FinalPrefix + $FinalText + $FinalSuffix

    #Skip separators if specified
    if ($NoSeparators)
    {
        #Don't add them
    }
    else
    {
        $FinalText = $StartSeparator + $FinalText + $EndSeparator
    }

    $Body = @{
        "content" = $FinalText
    }
    
    #Convert hashtable to json
    $BodyAsJson = ConvertTo-Json $Body
    
    #Send the text to the webhook
    Invoke-RestMethod -ContentType "application/json" -Body $BodyAsJson -Method Post -URI $FinalWebhook
}

<#
.synopsis
    * Plays a notification beep
.description
    * *BEEEEEEEEP*
.parameter Pitch
    * 2000 is kinda high and audible. The default.
    * Nothing below 190
.parameter Length
    * 500 is reasonable. The default.
.notes
    * Author: Ben Renninson
    * Email: ben@goldensyrupgames.com
    * From: https://github.com/GSGBen/Posh-Notify
#>
function Beep
{
    Param
    (
        [Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true)][int]$Pitch = 2000,
        [Parameter(Position=1,Mandatory=$false)][int]$Length = 500
    )

    [console]::beep($Pitch,$Length)
}

<#
.synopsis
    * Plays a notification Arpeggio
.description
    * *BEEEEEEEEP*
.notes
    * Author: Ben Renninson
    * Email: ben@goldensyrupgames.com
    * From: https://github.com/GSGBen/Posh-Notify
#>
function Arpeggio
{
    Param
    (
        [Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true)][int]$Pitch = 2000,
        [Parameter(Position=1,Mandatory=$false)][int]$Length = 500
    )

    Beep 880     300
    Beep 1108.73 300
    Beep 1318.51 300
    Beep 1760    300
}     

<#
.synopsis
    * Plays a built-in windows success sound
.description
    * *TADA*
.notes
    * Author: Ben Renninson
    * Email: ben@goldensyrupgames.com
    * From: https://github.com/GSGBen/Posh-Notify
#>
function Tada
{

    $PlayWav = New-Object System.Media.SoundPlayer
    $PlayWav.SoundLocation = "C:\Windows\media\tada.wav"
    $PlayWav.playsync()

}     


#region----------ALIASES

    New-Alias -Name Notify-Office365 -Value Send-Office365Notification
    New-Alias -Name Notify-Teams -Value Send-TeamsNotification
    New-Alias -Name Notify-Outlook -Value Send-OutlookNotification

    New-Alias -Name Notify-Discord -Value Send-DiscordNotification
    
    Export-ModuleMember -Alias *

#endregion