' Check command line parameters
Select Case WScript.Arguments.Count
	Case 0
		' Default if none specified is local computer (".")
		Set objWMIService = GetObject( "winmgmts://./root/cimv2" )
		Set colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )
		For Each objItem in colItems
			strComputer = objItem.Name
		Next
	Case 1
		' Command line parameter can either be a computer
		' name or "/?" to request online help
		strComputer = UCase( Wscript.Arguments(0) )
		if InStr( strComputer, "?" ) > 0 Then Syntax
	Case Else
		' Maximum is 1 command line parameter
		Syntax
End Select

' Define constants
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

' Header line for screen output
strMsg = vbCrLf & "CPU load percentage for " & strComputer & ":" & vbCrLf & vbCrLf

' Enable error handling
On Error Resume Next

' Connect to specified computer
Set objWMIService = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
' Display error number and description if applicable
If Err Then ShowError

' Query processor properties
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)
' Display error number and description if applicable
If Err Then ShowError
' Prepare display of results
For Each objItem In colItems
	strMsg = strMsg _
	       & "Device ID       : " & objItem.DeviceID       & vbCrLf _
	       & "Load Percentage : " & objItem.LoadPercentage & vbCrLf & vbCrLf
Next

' Display results
WScript.Echo strMsg

'Done
WScript.Quit(0)


Sub ShowError()
	strMsg = vbCrLf & "Error # " & Err.Number & vbCrLf & _
	         Err.Description & vbCrLf & vbCrLf
	Syntax
End Sub


Sub Syntax()
	strMsg = strMsg & vbCrLf _
	       & "CPULoad.vbs,  Version 1.00" & vbCrLf _
	       & "Display CPU load percentage for each processor " _
	       & "on any computer on the network" & vbCrLf & vbCrLf _
	       & "Usage:  CSCRIPT  //NOLOGO  CPULOAD.VBS  " _
	       & "[ computer_name ]" & vbCrLf & vbCrLf _
	       & "Where:  " & Chr(34) & "computer_name" & Chr(34) _
	       & " is the optional name of a remote" & vbCrLf _
	       & "        computer (default is local computer " _
	       & "name)" & vbCrLf & vbCrLf _
	       & "Written by Rob van der Woude" & vbCrLf _
	       & "http://www.robvanderwoude.com" & vbCrLf & vbCrLf _
	       & "Created with Microsoft's Scriptomatic 2.0 tool" & vbCrLf _
	       & "http://www.microsoft.com/downloads/details.aspx?" & vbCrLf _
	       & "    FamilyID=09dfc342-648b-4119-b7eb-783b0f7d1178&DisplayLang=en" & vbCrLf
	WScript.Echo strMsg
	WScript.Quit(1)
End Sub
