#requires -version 2

# (c) 2010 by John Robbins\Wintellect – Do whatever you want to do with it 
# as long as you give credit. 

<#.SYNOPSIS 
PowerShellPoint is the *only* way to do a presentation on PowerShell. All 
PowerShell, all the time!
.DESCRIPTION 
If you're doing a presentation on using PowerShell, there's Jeffrey Snover's 
excellent Start-Demo, (updated by Joel Bennett (http://poshcode.org/705)) for 
running the actual commands. However, to show discussion and bullet points, 
everyone switches to PowerPoint. That's crazy! EVERYTHING should be done in 
PowerShell when presenting PowerShell. Hence, PowerShellPoint!

To create your "slides" the format is as follows:
Slide titles start with an exclamation point.
Comment (#) are ignored. 
The slide points respect any white space and blank lines you have.
All titles and slide points are indented one character.

Here's an example slide file:
------
# A comment line that's ignored.
# An exclamation point on it's own is treated as a title slide.
# and the title area is not shown.
!
   <Title Goes Here>
!First Slide Title
Point 1
    Sub Point A
Point 2
    Sub Point B
!Second Slide Title
Point 3
    Sub Point C
Point 4
    Sub Point D
!Third Slide Title
Point 5
    Sub Point E
------

The script will validate that no slides contain more points than can be 
displayed or individual points will wrap.

The default is to switch the window to 78 x 24 but you can specify the window size
as parameters to the script.

The script properly saves and restores the original screen size and buffer on
exit.

When presenting with PowerShellPoint, use the 'h' command to get help.

For a typical presentation your screen is 1024 x 768. To make the text as large as 
possible, set the font to Consolas 28 point and the PowerShell window will fill the
screen. You should create a special shortcut and set the fonts in there so you don't 
mess up your existing PowerShell shortcuts.

.PARAMETER File 
The file that contains your slides. Defaults to .\Slides.txt.
.PARAMETER Width
The width in characters to make the screen and buffer. Defaults to 78.
.PARAMETER Height
The height in characters to make the screen and bugger. Defaults to 24.
.PARAMETER TitleForeground
The foreground color for the title and footer. Defaults to Yellow.
.PARAMETER TitleBackground
The background color for the title and footer. Defaults to Black.
.PARAMETER TextForeGround
The foreground color for the slide content text. Defaults to the current host foreground color.
.PARAMETER TextBackground
The background color for the slide content text. Defaults to the current host background color.
#>

param( [string]$File = ".\Slides.txt",
       [int]$Width   = 78,
       [int]$Height  = 24,
       # The foreground and background colors for the title and footer text.
       $TitleForeground = "Yellow",
       $TitleBackground = "Black",
       # Slide points foreground and background.
       $TextForeGround  = $Host.UI.RawUI.ForegroundColor,
       $TextBackGround  = $Host.UI.RawUI.BackgroundColor)

Set-StrictMode –version Latest

<# Versions
1.1 - Fixed : No longer redraws the screen if you press 'p' on the first
              slide.
      Added : Can specify intro slides with no title by using ! followed
              by no text. Slides with this format will not have the three
              line band drawn at the top of the screen.
      Fixed : Moved the colors from hard coded values to parameters.
      Fixed : Check to not allow running in the ISE.
1.0 - Initial Release
#>

$scriptVersion = "1.1"

# A function for reading in a character swiped from Jaykul's 
# excellect Start-Demo 3.3.3.
function Read-Char() 
{
  $inChar=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
  # loop until they press a character, so Shift or Ctrl, etc don't terminate us
  while($inChar.Character -eq 0)
  {
    $inChar=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
  }
  return $inChar.Character
}

function ProcessSlides($inputFile)
{
    $rawLines = Get-Content $inputFile
    
    # Contains the actual slides. The key is the slide number and the value are the
    # text lines.
    $slides = @{}

    # The slide number I'm processing.
    $slideNumber = 0
    [string[]]$lines = $null 

    # Process the raw text by reading it into the hash table.
    for ($i = 0 ; $i -lt $rawLines.Count ; $i++ )
    {
        # Skip any comment lines.
        if ($rawLines[$i].Trim(" ").StartsWith("#"))
        {
            continue
        }
        
        # Lines starting with "!" are a title.
        if ($rawLines[$i].StartsWith("!"))
        {
            if ($lines -ne $null)
            {
                $slides.Add($slideNumber,$lines)
                $lines = $null        
            }
            $slideNumber++ 
            $lines += $rawLines[$i].Trim(" ").Substring(1)
        }
        else
        {
            if ($slideNumber -eq 0)
            {
                throw "The first line must be a title slide starting with !"
            }
            
            # Make sure the line won't wrap.
            if ($rawLines[$i].Length -gt ($Width - 1))
            {
                Write-Warning "Slide line: $rawLines[$i] is too wide for the screen" -WarningAction Inquire
            }
            
            $lines += $rawLines[$i]
            
            # Check to see if this slide is bigger than the height
            if ($lines.Length -gt ($Height - 4))
            {
                $title = $lines[0]
                Write-Warning "Slide $title is too long for the screen" -WarningAction Inquire
            }
        }
    }
    
    # Add the last slide.
    $slides.Add($slideNumber,$lines)

    # Do some basic sanity checks on the slides.
    if ($slides.Keys.Count -eq 0)
    {
        throw "Input file '$File' does not look properly formatted."
    }
    return $slides
}

function Draw-Title($title)
{
    $cursorPos = $Host.UI.RawUI.CursorPosition
    $cursorPos.x = 0
    $cursorPos.y = 0
    $Host.UI.RawUI.CursorPosition = $cursorPos
    
    # If the slide title is empty, this is a title slide so don't do the title.
    if ($title -ne "")
    {
        Write-Host -NoNewline -back $TitleBackground -fore $TitleForeground $(" " * $Width)
        Write-Host -NoNewline -back $TitleBackground -fore $TitleForeground " " $title $(" " * ($Width - $title.Length - 3))
        Write-Host -NoNewline -back $TitleBackground -fore $TitleForeground $(" " * $Width)
    }
}

function Draw-SlideText($lines)
{
    $cursorPos = $Host.UI.RawUI.CursorPosition
    $cursorPos.x = 0
    $cursorPos.y = 4
    $Host.UI.RawUI.CursorPosition = $cursorPos
    
    for ($i = 1 ; $i -lt $lines.Count ; $i++ )
    {
        Write-Host " " $lines[$i]
    }
}

function Draw-Footer($slideNumber,$slideCount)
{
    $cursorPos = $Host.UI.RawUI.CursorPosition
    $cursorPos.y = $Height - 1
    $cursorPos.x = 0
    $Host.UI.RawUI.CursorPosition = $cursorPos
    
    $footer = "$slideNumber of $slideCount"
    Write-Host -NoNewline -back $TitleBackground -fore $TitleForeground "$(" " * ($Width - $footer.Length - 2)) $footer"
}

function Draw-BackScreen($message)
{
    $cursorPos = $Host.UI.RawUI.CursorPosition
    $cursorPos.x = 0
    $cursorPos.y = 0
    $Host.UI.RawUI.CursorPosition = $cursorPos
    
    $spaces = $(" " * $Width)
    for ($i = 0 ; $i -lt $Height ; $i++)
    {
        Write-Host -NoNewline -BackgroundColor black -ForegroundColor yellow $spaces
    }
    
    $cursorPos.x = ($Width / 2) - ($message.Length / 2)
    $cursorPos.y = 0
    $Host.UI.RawUI.CursorPosition = $cursorPos
    
    Write-Host -NoNewline -BackgroundColor black -ForegroundColor White $message
}

function Show-UsageHelp
{
    $help = @"
    
PowerShellPoint Help $scriptVersion - John Robbins - john@wintellect.com

Key             Action
---             ------
'n', '<space>'  Next slide
'p'             Previous slide
's'             Shell out to PowerShell
'h', '?'        This help
'q'             Quit

Press any key now to return to the current slide.
"@
    $cursorPos = $Host.UI.RawUI.CursorPosition
    $cursorPos.x = 0
    $cursorPos.y = 0
    $Host.UI.RawUI.CursorPosition = $cursorPos
    
    $spaces = $(" " * $Width)
    for ($i = 0 ; $i -lt $Height ; $i++)
    {
        Write-Host -NoNewline -BackgroundColor black -ForegroundColor yellow $spaces
    }
    
    $cursorPos.x = 0
    $cursorPos.y = 0
    $Host.UI.RawUI.CursorPosition = $cursorPos
    
    Write-Host -NoNewline -BackgroundColor black -ForegroundColor White $help
  
}

function main 
{
    # Check to see that we weren't started in the ISE. This probably needs to 
    # be internationalized.
    if ($Host.Name -eq "Windows PowerShell ISE Host")
    {
        Write-Warning "PowerShellPoint can only run in a console window. Sorry."
        return
    }

    # Save off the original window data.
    $originalWindowSize = $Host.UI.RawUI.WindowSize
    $originalBufferSize = $Host.UI.RawUI.BufferSize
    $originalTitle      = $Host.UI.RawUI.WindowTitle
    $originalBackground = $Host.UI.RawUI.BackgroundColor
    $originalForeground = $Host.UI.RawUI.ForegroundColor

    # Make sure the file exists. If not, give the user a chance to 
    # enter it.
    $File = Resolve-Path $File
    while(-not(Test-Path $File)) 
    {
        $File = Read-Host "Please enter the path of your slides file (Crtl+C to cancel)"
        $File = Resolve-Path $File
    }
    
    try
    {
        # Set the new window and buffer sizes to be the same so 
        # there are no scroll bars.
        $scriptWindowSize = $originalWindowSize
        $scriptWindowSize.Width = $Width
        $scriptWindowSize.Height = $Height
        $scriptBufferSize = $scriptWindowSize
        
        $Host.UI.RawUI.BackgroundColor = $TextBackGround
        $Host.UI.RawUI.ForegroundColor = $TextForeGround
        
        # Set the title.
        $Host.UI.RawUI.WindowTitle = "PowerShellPoint"

        # Read in the file and build the slides hash.
        $slides = ProcessSlides($File)

        # The slides are good to go so now resize the window.
        $Host.UI.RawUI.WindowSize = $scriptWindowSize
        $Host.UI.RawUI.BufferSize = $scriptBufferSize
        
        # Keeps track of the slide we are on.
        [int]$currentSlideNumber = 1
        # The flag to break out of displaying slides.
        [boolean]$keepShowing = $true
        # Flag to avoid redrawing the screen for unknown keypresses.
        [boolean]$redrawScreen = $true
                
        do
        {
            if ($redrawScreen -eq $true)
            {
                Clear-Host 
                
                # Grab the current slide.
                $slideData = $slides.$currentSlideNumber
                
                Draw-Title $slideData[0]
                Draw-SlideText $slideData
                Draw-Footer $currentSlideNumber $slides.Keys.Count
            }
            
            $char = Read-Char
            
            switch -regex ($char)
            {
                # Next slide processing.
                "[ ]|n"
                {
                    $redrawScreen = $true
                    $currentSlideNumber++

                    if ($currentSlideNumber -eq ($slides.Keys.Count + 1))
                    {
                        # Pretend you're PowerPoint and show the black screen
                        Draw-BackScreen "End of slide show"
                        $ch = Read-Char
                        if ($ch -eq "p")
                        {
                            $currentSlideNumber--    
                        }
                        else
                        {
                            $keepShowing = $false
                        }
                    }
                }
                # Previous slide processing.
                "p"
                {
                    $redrawScreen = $true
                    $currentSlideNumber--

                    if($currentSlideNumber -eq 0)
                    {
                        $currentSlideNumber = 1
                        $redrawScreen = $false
                    }
                }
                # Quit processing.
                "q"
                {
                    $keepShowing = $false
                }
                "s"
                {
                    Clear-Host
                    Write-Host -ForegroundColor $TitleForeground -BackgroundColor $TitleBackground "Suspending PowerShellPoint - type 'Exit' to resume"
			        $Host.EnterNestedPrompt()
                }
                # Help processing.
                "h|\?"
                {
                    Show-UsageHelp
                    $redrawScreen = $true
                    Read-Char
                }
                # All other keys fall here.
                default
                {
                    $redrawScreen = $false
                }
            }
        } while ($keepShowing)


        # The script has finished cleanly so clear the screen.
        $Host.UI.RawUI.BackgroundColor = $originalBackground
        $Host.UI.RawUI.ForegroundColor = $originalForeground 
        Clear-Host
    }    
    finally
    {
        # I learned something here. You have to set the buffer size before 
        # you set the window size or the window won't resize.
        $Host.UI.RawUI.BufferSize  = $originalBufferSize
        $Host.UI.RawUI.WindowSize  = $originalWindowSize 
        $Host.UI.RawUI.WindowTitle = $originalTitle
        $Host.UI.RawUI.BackgroundColor = $originalBackground
        $Host.UI.RawUI.ForegroundColor = $originalForeground 
    }
}

. main
# SIG # Begin signature block
# MIITMQYJKoZIhvcNAQcCoIITIjCCEx4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmbvOZXA182mgK6UqrMDGsbqt
# ug+ggg5AMIIEhTCCA22gAwIBAgIDAjpkMA0GCSqGSIb3DQEBBQUAMEIxCzAJBgNV
# BAYTAlVTMRYwFAYDVQQKEw1HZW9UcnVzdCBJbmMuMRswGQYDVQQDExJHZW9UcnVz
# dCBHbG9iYWwgQ0EwHhcNMTIxMDE4MTQzODM1WhcNMjIwNTIwMTQzODM1WjBeMQsw
# CQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNV
# BAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4
# JsRDc2vCvy5QWvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9F
# O9FEzkMScxeCi2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M
# 4lc/PcaS3Er4ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHw
# Hw103pIiq8r3+3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcs
# n6plINPYlujIfKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEA
# AaOCAWYwggFiMB8GA1UdIwQYMBaAFMB6mGiNifurBWQMEX2qfWW4ysxOMB0GA1Ud
# DgQWBBRfmvVuXMzMdJrU3X3vP9vsTIAu3TASBgNVHRMBAf8ECDAGAQH/AgEAMA4G
# A1UdDwEB/wQEAwIBBjA6BgNVHR8EMzAxMC+gLaArhilodHRwOi8vY3JsLmdlb3Ry
# dXN0LmNvbS9jcmxzL2d0Z2xvYmFsLmNybDA0BggrBgEFBQcBAQQoMCYwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmdlb3RydXN0LmNvbTBLBgNVHSAERDBCMEAGCSsG
# AQQB8CIBBzAzMDEGCCsGAQUFBwIBFiVodHRwOi8vd3d3Lmdlb3RydXN0LmNvbS9y
# ZXNvdXJjZXMvY3BzMCgGA1UdEQQhMB+kHTAbMRkwFwYDVQQDExBUaW1lU3RhbXAt
# MjA0OC0xMBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEBBQUAA4IBAQCq
# Pj8DQS81E3+MPxJcEM/uZ1V5VmAPAVIPI4d6QShH5dNmcrsCz8kMJClBwtoumeAj
# ke+xM0c8wgg6BpVzohsOvD+34lnxab7w0+EYj3Mt6KOcIUBA8806twGN3E2UtHoQ
# BVB/G2HFghWK5CxN7TQR61tqiVnH3vcshMCzvTqY4UUoiiGVgD/JB5fw/0LBHkKE
# 57LH4KJqmdTx1Mb+V8C5Ouf2J3ANqeB7ShOVwsm5q2kH2U3MZkdCSnSnp22zpzXo
# fkuqzr28k599xkCl/KofqNXR8H5/0uNMVv1AZPBHTFEEMOrqlC8kyukIxXhboWLJ
# HUiHlxn6TaV47ewEeQ42MIIEozCCA4ugAwIBAgIQfh/fcpno0kWhXQuo5bFZujAN
# BgkqhkiG9w0BAQUFADBeMQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMg
# Q29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2Vy
# dmljZXMgQ0EgLSBHMjAeFw0xMjEwMTgwMDAwMDBaFw0yMjA1MTkyMzU5NTlaMGIx
# CzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjE0MDIG
# A1UEAxMrU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBTaWduZXIgLSBH
# NDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKJjCzlEuLsjp0RJuw7/
# ofBhClOTsJjbrSwPSsVu/4Y8U1UPFc4EPyv9qZaW2b5heQtbyUyGduXgQ0sile7C
# K0PBn9hotI5AT+6FOLkRxSPyZFjwFTJvTlehroikAtcqHs1L4d1j1ReJMluwXpla
# qJ0oUA4X7pbbYTtFUR3PElYLkkf8q672Zj1HrHBy55LnX80QucSDZJQZvSWA4ejS
# IqXQugJ6oXeTW2XD7hd0vEGGKtwITIySjJEtnndEH2jWqHR32w5bMotWizO92WPI
# SZ06xcXqMwvS8aMb9Iu+2bNXizveBKd6IrIkri7HcMW+ToMmCPsLvalPmQjhEChy
# qs0CAwEAAaOCAVcwggFTMAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwDgYDVR0PAQH/BAQDAgeAMHMGCCsGAQUFBwEBBGcwZTAqBggrBgEFBQcw
# AYYeaHR0cDovL3RzLW9jc3Aud3Muc3ltYW50ZWMuY29tMDcGCCsGAQUFBzAChito
# dHRwOi8vdHMtYWlhLndzLnN5bWFudGVjLmNvbS90c3MtY2EtZzIuY2VyMDwGA1Ud
# HwQ1MDMwMaAvoC2GK2h0dHA6Ly90cy1jcmwud3Muc3ltYW50ZWMuY29tL3Rzcy1j
# YS1nMi5jcmwwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0yMDQ4
# LTIwHQYDVR0OBBYEFEbGaaMOShQe1UzaUmMXP142vA3mMB8GA1UdIwQYMBaAFF+a
# 9W5czMx0mtTdfe8/2+xMgC7dMA0GCSqGSIb3DQEBBQUAA4IBAQBjADKNFx8o3AR5
# yScZhLg2aKd1GRiyT0msXWErhjLQDU26tXxettI36O1biCzSKWG+H1ApSiL5F4a9
# hyFb0TxNv2TAui6bphCh8cSEU7CNWPKOrxIZdx+t976+gS2Ogn5w+DmWM2VZqE9/
# iyLJGH5eZOK5MG0GtLcRjGa6LCZEuYrcsYeRtdy/FKHcg6NgryleZoorDe2d0DkF
# ubhvHrm6ccxo3rJ5OgPWiOEdKbM05yHYow8rchpCtZ5F+t+nK53X8s0dyFYSL42d
# Tc4yySZhZNCaeOUvEmpKRAIk2FeFMn/NWYxrczMApV6rw/AiwMsIw+D3uLgEFOxK
# 7zn5z8olMIIFDDCCA/SgAwIBAgIQP/vU6E1XgR79hivMOYXcWzANBgkqhkiG9w0B
# AQUFADCBlTELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlVUMRcwFQYDVQQHEw5TYWx0
# IExha2UgQ2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMSEwHwYD
# VQQLExhodHRwOi8vd3d3LnVzZXJ0cnVzdC5jb20xHTAbBgNVBAMTFFVUTi1VU0VS
# Rmlyc3QtT2JqZWN0MB4XDTEwMTExNzAwMDAwMFoXDTEzMTExNjIzNTk1OVowgZ0x
# CzAJBgNVBAYTAlVTMQ4wDAYDVQQRDAUzNzkzMjELMAkGA1UECAwCVE4xEjAQBgNV
# BAcMCUtub3h2aWxsZTESMBAGA1UECQwJU3VpdGUgMzAyMR8wHQYDVQQJDBYxMDIw
# NyBUZWNobm9sb2d5IERyaXZlMRMwEQYDVQQKDApXaW50ZWxsZWN0MRMwEQYDVQQD
# DApXaW50ZWxsZWN0MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEApF66
# GIwwpYHMG9CAW8yTT4Mb5mw/af8tyrBZpTglGkrnxAhIdJyrVajft7H3ZV8rWLgc
# WT+JQ7BG8H6daaQX3npiIl1g4dGUGBCWKrt/2Fqq2DcjCSnnic+ykbzfKVeDdrvp
# wCAz1SRX40qgNe2WEVWbPBYoEQ4HGCaZwBpAQ4yFkypAcHPzYZPcNVyoQxX2cL75
# 2HgpzWZLz8bJRDrv+aeVliVtJ1qb0/QWP2/T0fyQ0bHZxa5b2bs65OzQ8jbwzHMy
# glM0pxx2YetLbUVdd0WqtIA2irEq3D5OqQDfx7DyUAFEDQVj/twVplemBX0Gwnyp
# mWQReEn/6uUnV7AwOwIDAQABo4IBTDCCAUgwHwYDVR0jBBgwFoAU2u1kdBScFDyr
# 3ZmpvVsoTYs8ydgwHQYDVR0OBBYEFOamMI47Dp8ULxWaFqGX6erMuF7iMA4GA1Ud
# DwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMBEG
# CWCGSAGG+EIBAQQEAwIEEDBGBgNVHSAEPzA9MDsGDCsGAQQBsjEBAgEDAjArMCkG
# CCsGAQUFBwIBFh1odHRwczovL3NlY3VyZS5jb21vZG8ubmV0L0NQUzBCBgNVHR8E
# OzA5MDegNaAzhjFodHRwOi8vY3JsLnVzZXJ0cnVzdC5jb20vVVROLVVTRVJGaXJz
# dC1PYmplY3QuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDov
# L29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBBQUAA4IBAQBIft0bPwDq/yHv
# 6m58y4PeMh3yIqcV+89TgsoVu2DUzgrOPBCF182icb0A+ezvktx6BwdmRgNXGaJJ
# YuqAEDZy3j+dZMUCSnE1ci9rEpCklf9nr2LSJoUU0QmTF3tDqxnS90cu8fRGmoka
# 4cH1SQRm6nnlnOnKaxRpJA5uyqSiqAsY7wxd35ENVdEBCV0C05xJ3UAfnCyiJT/r
# MFkEMWC4X4Pj/LUPqsAG5wTWGe/i+U4sGfD6dEuvk/eEorqBvORNaGkFLTU9UjvB
# S0GuDN2zucXRgNhi9VbYv1vBqfLSD6/8DAJOF7fvnprIneUxw2ZI0qxTKmwh5msa
# E5rXLra+MYIEWzCCBFcCAQEwgaowgZUxCzAJBgNVBAYTAlVTMQswCQYDVQQIEwJV
# VDEXMBUGA1UEBxMOU2FsdCBMYWtlIENpdHkxHjAcBgNVBAoTFVRoZSBVU0VSVFJV
# U1QgTmV0d29yazEhMB8GA1UECxMYaHR0cDovL3d3dy51c2VydHJ1c3QuY29tMR0w
# GwYDVQQDExRVVE4tVVNFUkZpcnN0LU9iamVjdAIQP/vU6E1XgR79hivMOYXcWzAJ
# BgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAj
# BgkqhkiG9w0BCQQxFgQUMFQp/RTZXJGza7BlyskE15dhyBkwDQYJKoZIhvcNAQEB
# BQAEggEAEUEVmrU/AT1F0dnZrkgYAIi6m9qt/CIA4/XZgD9qSUTCfxbDqOsSQh5+
# Fa+JxM0hI6SwqeSKhiP9R1UuuV/RbAoGNpghDHCyc1DVF8JNIXqW8Xt+DvZK+6Vg
# oxgepzPMYhVsbCAuQtMOorNSjaCCZ9gSAjRYlZLb7t9Xr53L1+LKoLjC6yDJBMuS
# v97P8786uGtNOVuaBGzZBwrPo/OSGHAlMfW+eCrYcb/90bpKTvhKTzbTsQuRzawd
# xVqzMDTSP3LOaTq69GmcTVKSUnuL6js2abY9xYnE/CmQKgo7nkFCNBrulYYhrFt3
# XYve05P6p84Im6zLaEhI70lLpPEz+KGCAgswggIHBgkqhkiG9w0BCQYxggH4MIIB
# 9AIBATByMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyAhB+H99ymejSRaFdC6jlsVm6MAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0B
# CQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xMjEyMjEyMTQ0MTRaMCMG
# CSqGSIb3DQEJBDEWBBQMPfIUZ/+HsXcjloy2sm/AsC+HwDANBgkqhkiG9w0BAQEF
# AASCAQBrshJ2IhfYhRzwvMZ7hS5c+Nqpp6rIvSvyP3hhhoV9AI9HA14F7lFWVMII
# ExMe/m4TLwb0PnOUgniO9f7BTOZr4DbIsaDWOGLn2iAZznpbcEyqlRaJqYuqJ3CD
# Y+oBN7PtbN0SAKyKycBwtCM6ynzNqYmZ57xlobBC25hlgtE/qUpKTVXjtcbV2zHT
# HdP09yxJsn15B57RfXrcJvhm+rmuJsHQ3NUqkbz6zBnmfZCw8HdW2xSsJ//LC53D
# m7WyQvkMJjgkDK0ED0A2O/bvcQvNYTbEan6sh0B7PY0ZT/w2opnOL4o6VwkYIucS
# b+WMwGzRCfrgz+BE99wl3+jXOOPY
# SIG # End signature block
