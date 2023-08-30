<#
.SYNOPSIS 
    PowerShell script to set your Teams status to your Last.fm now playing track

.DESCRIPTION
    PowerShell script to set your Teams status to your Last.fm now playing track.
    This script uses the Microsoft.Graph.Beta module to connect to the Microsoft Graph API and set your status.
    It will check every 10 seconds if you're scrobbling anything on Last.fm and set your status accordingly.
    If you're not scrobbling anything, it will clear your status message (if it was set by this script).

.EXAMPLE
    # Set your Last.fm API key
    $env:LAST_FM_API_KEY = "your_api_key_here"

    # Run the script
    .\LastFmStatus.ps1

    # Press Ctrl+C to stop the script

.LINK
    https://github.com/plusreed/MsTeamsLastFmStatus
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$LastFmUser,
    [Parameter(Mandatory = $false)]
    [int32]$SleepTime = 10
)

Import-Module Microsoft.Graph.Beta.Users.Actions

# Check if env var is set
if ($nul -eq $env:LAST_FM_API_KEY) {
    Write-Error "LAST_FM_API_KEY environment variable not set. Please set it to your Last.fm API key."
    exit 1
}

$consts = @{
    # how often to check for a new track (in seconds)
    sleepTime        = $SleepTime

    apiUrl           = "https://ws.audioscrobbler.com/2.0/?method=user.getrecenttracks&user={0}&api_key={1}&format=json&limit=1"

    # {0} = track name
    # {1} = artist name
    # {2} = album name
    tmplLastFmStatus = "$([System.Char]::ConvertFromUtf32(127926)) Now playing:<br>{1} - {0} [{2}]"

    # your last.fm username, passed in as -LastFmUser
    lastFmUser       = $LastFmUser
}
$global:lastStatus = $null
$global:previousEncoding = [Console]::OutputEncoding

# set encoding to utf8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Function Get-CurrentTime {
    return (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
}

Write-Host "Last.fm Status for Microsoft Teams starting..."
Write-Host "Started at $(Get-CurrentTime)"
Write-Host "$('=' * 60)"

Write-Host "$(Get-CurrentTime)`tLogs start here!"
Write-Host "$(Get-CurrentTime)`tConnecting to Microsoft Graph..."

Connect-MgGraph -Scopes "User.ReadWrite,Presence.ReadWrite" -NoWelcome

# Get my profile
Write-Debug "$(Get-CurrentTime)`tGetting user ID..."
$me = Invoke-MgRestMethod -Uri '/v1.0/me'
$meId = $me['id']

Write-Debug "$(Get-CurrentTime)`tyour user ID: $meId"

Function Get-LastFmNowPlaying {
    $lastFmApiReq = Invoke-WebRequest -Uri ($consts.apiUrl -f $consts.lastFmUser, $env:LAST_FM_API_KEY)
    $lastfmApiReqContent = [System.Text.Encoding]::UTF8.GetString($lastFmApiReq.Content.ToCharArray()) | ConvertFrom-Json

    $track = $lastFmApiReqContent.recenttracks.track[0]

    # is this a now playing track?
    if ($track.'@attr'.nowplaying -eq $true) {
        # get the track info
        $trackName = $track.name
        $artistName = $track.artist.'#text'
        $albumName = $track.album.'#text'
        
        return @{
            trackName  = $trackName
            artistName = $artistName
            albumName  = $albumName
        }
    }
    else {
        return $null
    }
}

Function FunLittleLoopyThing {
    # get the now playing status from last.fm
    $lastFmStatus = Get-LastFmNowPlaying

    # is this a now playing track?
    if ($null -ne $lastFmStatus) {
        # set the status
        $status = $consts.tmplLastFmStatus -f $lastFmStatus.trackName, $lastFmStatus.artistName, $lastFmStatus.albumName

        # Check if the generated status differs from the last one
        if ($status -eq $global:lastStatus) {
            Write-Host "$(Get-CurrentTime)`tStatus hasn't changed, skipping this loop..."
            return
        }

        Write-Host "$(Get-CurrentTime)`tSetting status to: $status"
        $global:lastStatus = $status

        $statusMessageParams = @{
            statusMessage = @{
                message        = @{
                    content     = $status
                    contentType = "text"
                }
                expiryDateTime = @{
                    dateTime = (Get-Date).AddHours(1).ToString("yyyy-MM-ddTHH:mm:ssZ")
                    timeZone = "Eastern Standard Time"
                }
            }
        }

        Set-MgBetaUserPresenceStatusMessage -UserId $meId -BodyParameter $statusMessageParams
    }
    else {
        if ($null -eq $global:lastStatus) {
            Write-Host "$(Get-CurrentTime)`tNot scrobbling anything (no status message set)..."
            return
        }
        else {
            # Clear our existing status message
            $statusMessageParams = @{
                statusMessage = @{
                    message = @{
                        content     = ""
                        contentType = "text"
                    }
                }
            }
            Set-MgBetaUserPresenceStatusMessage -UserId $meId -BodyParameter $statusMessageParams
            Write-Host "$(Get-CurrentTime)`tNot scrobbling anything (clearing status message)..."
            $global:lastStatus = $null
        }
    }

    return
}
try {
    while ($true) {
        FunLittleLoopyThing

        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                Write-Host "$(Get-CurrentTime)`tDisconnecting from Microsoft Graph..."
                Disconnect-MgGraph
                break
            }
        }

        Write-Debug "$(Get-CurrentTime)`tsleeping for $($consts.sleepTime) seconds..."
        Start-Sleep -Seconds 10
    }
}
finally {
    # disconnect from graph
    Write-Host "$(Get-CurrentTime)`tDisconnecting from Microsoft Graph..."
    Disconnect-MgGraph

    Write-Host "$(Get-CurrentTime)`tStopped at $(Get-CurrentTime)"
    Write-Host "$('=' * 60)"

    # reset encoding
    [Console]::OutputEncoding = $global:previousEncoding
}