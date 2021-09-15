##############################################Set-ScriptLocation
<#
.SYNOPSIS
Set script directory as working directory

.DESCRIPTION
Set script directory as working directory

.EXAMPLE
Set-ScriptLocation

.NOTES
General notes
#>
function Set-ScriptLocation {
  Set-Location $MyInvocation.PSScriptRoot
}
function Get-ScriptLocation {
  return $MyInvocation.PSScriptRoot
}
##############################################Set-FileShortcut
<#
.SYNOPSIS
Creates a shortcut to a file or folder

.DESCRIPTION
We use a call to VBS shell to create a windows .lnk file

.PARAMETER SourceFile
File to which the shortcut points

.PARAMETER DestinationPath
Path where the shortcut will be created

.PARAMETER OptionalLinkName
Optional name for the shortcut

.EXAMPLE
Set-FileShortcut -SourceFile "./main.ps1" -DestinationPath "./" -OptionalLinkName Test

.NOTES
General notes
#>
function Set-FileShortcut {
  # Ideas from https://stackoverflow.com/questions/9701840/how-to-create-a-shortcut-using-powershell
  param(
    $SourceFile,
    $DestinationPath,
    $OptionalLinkName
  )

  # Convert relative paths to absolute
  $SourceFile = (Resolve-Path $SourceFile).Path
  $DestinationPath = (Resolve-Path $DestinationPath).Path

  # Get the file basename
  $BaseName = (Get-ChildItem $SourceFile).BaseName

  # If optionalLinkName is set then
  if ($OptionalLinkName) {
    # LinkName equals OptionalLinkName
    $LinkName = $OptionalLinkName
  }
  else {
    # Else Linkname equals $BaseName
    $LinkName = $BaseName
  }
  # We spawn a new vbs shell
  $WshShell = New-Object -comObject WScript.Shell

  # We create the shortcut
  $Shortcut = $WshShell.CreateShortcut("$DestinationPath/$LinkName.lnk")

  # We define the target for the shortcut
  $Shortcut.TargetPath = $SourceFile

  # We save the shortcut
  $Shortcut.Save()

  Write-Host "$DestinationPath/$LinkName.lnk created" 
}

##############################################Send-ToEmail
<#
.SYNOPSIS
Sends and e-mail

.DESCRIPTION
This script sends an email by providing it the basic account configuration for SMTP

.PARAMETER Username
The username, usually the same as the sender e-mail

.PARAMETER Password
The sender email password

.PARAMETER Recipient
Which e-mail is going to receive the message

.PARAMETER SenderAddress
The sender e-mail, usually the same as the username

.PARAMETER Subject
Subject of the e-mail

.PARAMETER Body
The body of the e-mail

.PARAMETER Server
The e-mail server address (mail.example.net)

.PARAMETER Port
The SMTP port (25,465,587)

.PARAMETER EnableSSL
If you use port 465 or 587, set this to true

.EXAMPLE
Send-ToEmail -Username "admin@example.net" -SenderAddress "admin@example.net" -Password "rockyou" -Recipient "user@example.net" -Subject "Testing a cool PWSH library" -Body "So this is an example of how the body would look like, you can put the content of a variable on here as well" -Server "smtp.example.net" -Port 587 -EnableSSL $true

.NOTES
General notes
#>
function Send-ToEmail {
  param(
    [string]$Username,
    [securestring]$Password,
    [string]$Recipient,
    [string]$SenderAddress,
    [string]$Subject,
    [string]$Body,
    [string]$Server,
    [int16]$Port,
    [bool]$EnableSSL = $False
  )
  $message = new-object Net.Mail.MailMessage
  $message.From = $SenderAddress
  $message.To.Add($Recipient)
  $message.Subject = $Subject
  $message.Body = $Body
  $smtp = new-object Net.Mail.SmtpClient($Server, $Port)
  $smtp.EnableSSL = $EnableSSL
  $smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
  $smtp.send($message)
  write-host "Message Sent"
}

##############################################TimeStamp
<#
.SYNOPSIS
Gives you a timestamp which helps sort by date

.DESCRIPTION
Gives you a timestamp which helps sort by date

.EXAMPLE
(TimeStamp)

.NOTES
If you fill a variable with it, don't forget to update it.
#>
function TimeStamp {
  param(
    $format
  )
  if ($format -eq "file") {
    Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
  }
  elseif ($format -eq "log") {
    Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  }
  else {
    Write-Host 'Timestamp function: You need to define the format parameter to either "log" or "file"'
    break
  }

}

##############################################IsEmptyOrNull
<#
.SYNOPSIS
Checks if a variable or value is null or empty

.DESCRIPTION
Checks if a variable or value is null or empty

.PARAMETER evaluate
Pass to this parameter the value you want to evaluate

.EXAMPLE
IsEmptyOrNull -evaluate $null

.NOTES
$NULL
#>
function IsEmptyOrNull {
  param(
    $evaluate
  )
  [string]::IsNullOrWhiteSpace($evaluate)
}

<#
.SYNOPSIS
Allows to add logging to scripts

.DESCRIPTION
This function does logging to an external file in a way that felt confortable to me, probably can be improved but that's what I got for now.

.PARAMETER newline
The message you want to add to the log

.PARAMETER level
The level you want to show on the log line, by default Info, but can be "UberAssError", or whatever

.PARAMETER Path
The path for the log file, by default the base name of the script + .log

.EXAMPLE
Add-Log -newline "OhSh**" -level "Better run" -Path Example.log
gives:
2021-09-15 21:55:43 [Better run] OhSh**

.NOTES
Useful huh?
#>
function Add-Log {
  param(
    $newline,
    $level,
    $Path
  )

  # If path is not defined creates a file.log with the same name as the script name on the same path
  # BUG - WHEN INCLUDED USES LIBRARY FILE NAME - Fixed with $MyInvocation.PSCommandPath
  # ENHANCEMENT - If absolute path not defined, put file on same folder as script
  if (!($Path)) {
    # We rip the extension out of the script file name, and then we add the extension
    $logfile = ((Get-Item $MyInvocation.PSCommandPath).basename) + ".log"
    # We define the full path as current script path plus the filename we generated.
    $path = ($MyInvocation.PSScriptRoot) + "\$logfile"
  } # Else , if theres a path defined but the resolved path is not equal an absolute path, then is relative
  elseif ((Resolve-Path -ErrorAction Ignore $Path).Path -ne $Path) {
    # We set path as the script folder + the file name we specified.
    $path = ($MyInvocation.PSScriptRoot) + "\$Path"
  } # The condition left, is a full path, so we do nothing

  # If level not set, default level is info
  if (!$level) {
    $level = "Info"
  }
  # Log format ex: 
  #      2021-09-14 22:25:16           [Info]        This is dog
  # Depends on TimeStamp function to work
  $log = (TimeStamp -format "log") + " [$level]" + " $log" + "$newline"
  # Writing to file
  $log | Out-File -Append -FilePath $path
}
