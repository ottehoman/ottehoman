# simple mail message script
# requires COMMENT to be '-from "email" -subject "subject" -body "bodymsg" -attach "path\to\file"'
# default values are set below. If content in the given parameters, it will override that

# username/password are stored in a Generic Credential on the sending machine
# [CmdletBinding()] gives us some -debug and -verbose capabilities
[CmdletBinding()]
Param (
 $TO = "no-reply@unsw.edu.au",
 $SUBJECT = "FPL e-Line+ milestone :-)",
 $BODY = "Mission accomplished.",
 $ATTACH = "C:\Users\z3390106\OneDrive - UNSW\Documents\GitHub\ottehoman\eLine_Message.txt",
 $SERVER = "smtp.gmail.com",
 $PORT = 587,
 $FROM = "no-reply@unsw.edu.au",
 $CRED =$( Get-StoredCredential -Target $SERVER)
)

$StartTime = (Get-Date).Ticks

$RESULT = Send-MailMessage -smtpserver $SERVER  -Port $PORT -Credential $CRED -From $FROM -To $TO -Subject $SUBJECT -Body $BODY -Attachments $ATTACH -UseSsl -Verbose
Write-Host $RESULT


$StopTime = (Get-Date).Ticks
$DeltaTime = ($StopTime - $StartTime)
# one tick is 100ns

$SECS = [int]$DeltaTime/10000000
$SECSTR = " in {0:f2} seconds." -f ($SECS % 60)
$STR = "Email sent to " + $TO + " in " + ($SECS).ToString("#.#") + " seconds."

Write-Host $STR
