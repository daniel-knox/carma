Option Strict Off
Option Explicit On
Module RRKit_Common_Commands
	Public comDevice As Integer 'Global handle to RD/RS device connected
    Public Status As Integer    'Global status variable to communicate with DLL
    Public PhaseCount As Short  'Global phase count for device connected
    Public TestType As String   'Global test type for time/pulse test timer
    Public TapPhase As Byte     'Global variable for tap change phase selection

	'declarations for starting & stopping RadCommApp
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

	
	Sub RadCommApp_Connect()
		' ***************************************************************************
		' *                         Radian Research, Inc.                           *
		' *                  Copyright 2005  All Rights Reserved.                   *
		' ***************************************************************************
		' *                                                                         *
		' * PROJECT:    RD/RS Visual Basic Communications                           *
		' *                                                                         *
		' * FILE:       RRKit_Common_Commands.bas                                   *
		' * AUTHOR:     Mike Henderson                                              *
		' * DATE:       02/22/05                                                    *
		' *                                                                         *
		' * SUBROUTINE:  RadCommApp_Connect                                         *
		' *                                                                         *
		' * INPUTS: None                                                            *
		' *                                                                         *
		' * OUTPUTS: None                                                           *
		' *                                                                         *
		' * DESCRIPTION:  Starts RadCommApp if needed and sets up usage count       *
		' *                                                                         *
		' * CHANGE HISTORY:                                                         *
		' *                                                                         *
		' * COMPATABILITY: RD-2X/3X RS-1X RS-712                                    *
		' ***************************************************************************
		Dim commapp As Integer

		On Error GoTo EndError
		commapp = FindWindow(vbNullString, "RadCommApp")
		If commapp = 0 Then 'no radcommapp found in memory
			commapp = Shell("RadCommApp.exe")
		End If
		
Connected: 
		Call PostMessage(commapp, &H400s + 1, 0, 0)
		Exit Sub
		
EndError: 
		Call MsgBox("Could not find RadCommApp.exe.", MsgBoxStyle.Critical, "RadCommApp")
		End
	End Sub
	
	Sub RadCommApp_Release()
		' ***************************************************************************
		' *                         Radian Research, Inc.                           *
		' *                  Copyright 2005  All Rights Reserved.                   *
		' ***************************************************************************
		' *                                                                         *
		' * PROJECT:    RD/RS Visual Basic Communications                           *
		' *                                                                         *
		' * FILE:       RRKit_Common_Commands.bas                                   *
		' * AUTHOR:     Mike Henderson                                              *
		' * DATE:       02/22/05                                                    *
		' *                                                                         *
		' * SUBROUTINE:  RadCommApp_Release                                         *
		' *                                                                         *
		' * INPUTS: None                                                            *
		' *                                                                         *
		' * OUTPUTS: None                                                           *
		' *                                                                         *
		' * DESCRIPTION:  Releases usage count to RadCommApp                        *
		' *                                                                         *
		' * CHANGE HISTORY:                                                         *
		' *                                                                         *
		' * COMPATABILITY: RD-2X/3X RS-1X RS-712                                    *
		' ***************************************************************************
		Dim commapp As String
		
		On Error GoTo EndError
		If comDevice <> 0 Then RadRDReleaseDevice(comDevice)
		
		commapp = CStr(FindWindow(vbNullString, "RadCommApp"))
		If CDbl(commapp) <> 0 Then
			Call PostMessage(CInt(commapp), &H402s, 0, 0)
		End If
		Exit Sub
		
EndError: 
		Call MsgBox("Unable to properly release RadCommApp", MsgBoxStyle.Information, "RadCommApp")
	End Sub
	
    Function ConvertFromHex(ByVal Data As String) As String
        ' ***************************************************************************
        ' *                         Radian Research, Inc.                           *
        ' *                  Copyright 2005  All Rights Reserved.                   *
        ' ***************************************************************************
        ' *                                                                         *
        ' * PROJECT:    RD/RS Visual Basic Communications                           *
        ' *                                                                         *
        ' * FILE:       RRKit_Common_Commands.bas                                   *
        ' * AUTHOR:     Mike Henderson                                              *
        ' * DATE:       02/22/05                                                    *
        ' *                                                                         *
        ' * SUBROUTINE:  ConvertFromHex                                             *
        ' *                                                                         *
        ' * INPUTS: string Data - Hex string of data (i.e. A6030000 - Radian Reset) *
        ' *                                                                         *
        ' * OUTPUTS: string - ASCII string                                          *
        ' *                                                                         *
        ' * DESCRIPTION: Converts character string from hex representations to ASCII*
        ' *                                                                         *
        ' * CHANGE HISTORY:                                                         *
        ' *                                                                         *
        ' * COMPATABILITY: RD-2X/3X RS-1X RS-712                                    *
        ' ***************************************************************************
        Dim DefaultString As String
        Dim I As Long

        DefaultString = ""
        For I = 1 To Len(Data) Step 2
            DefaultString = DefaultString & Chr("&H" & Mid(Data, I, 2))
        Next I
        ConvertFromHex = DefaultString
    End Function

    Function ConverttoHex(ByVal Data As String) As String
        ' ***************************************************************************
        ' *                         Radian Research, Inc.                           *
        ' *                  Copyright 2005  All Rights Reserved.                   *
        ' ***************************************************************************
        ' *                                                                         *
        ' * PROJECT:    RD/RS Visual Basic Communications                           *
        ' *                                                                         *
        ' * FILE:       RRKit_Common_Commands.bas                                   *
        ' * AUTHOR:     Mike Henderson                                              *
        ' * DATE:       02/22/05                                                    *
        ' *                                                                         *
        ' * SUBROUTINE:  ConverttoHex                                               *
        ' *                                                                         *
        ' * INPUTS: string Data - ASCII string of data                              *
        ' *                                                                         *
        ' * OUTPUTS: string - Hex representation string                             *
        ' *                                                                         *
        ' * DESCRIPTION: Converts character string from ASCII to hex representations*
        ' *                                                                         *
        ' * CHANGE HISTORY:                                                         *
        ' *                                                                         *
        ' * COMPATABILITY: RD-2X/3X RS-1X RS-712                                    *
        ' ***************************************************************************
        Dim I As Long

        If Len(Data) = 0 Then
            ConverttoHex = ""
            Exit Function
        End If
        For I = 1 To Len(Data)
            ConverttoHex = ConverttoHex & Right("00" & Hex(Asc(Mid(Data, I, 1))), 2)
        Next I
    End Function

    Sub StatusError()
        ' ***************************************************************************
        ' *                         Radian Research, Inc.                           *
        ' *                  Copyright 2005  All Rights Reserved.                   *
        ' ***************************************************************************
        ' *                                                                         *
        ' * PROJECT:    RD/RS Visual Basic Communications                           *
        ' *                                                                         *
        ' * FILE:       RRKit_Common_Commands.bas                                   *
        ' * AUTHOR:     Mike Henderson                                              *
        ' * DATE:       02/22/05                                                    *
        ' *                                                                         *
        ' * SUBROUTINE:  StatusError                                                *
        ' *                                                                         *
        ' * INPUTS: None                                                            *
        ' *                                                                         *
        ' * OUTPUTS: None                                                           *
        ' *                                                                         *
        ' * DESCRIPTION: Take the status error number and display the error message *
        ' *                                                                         *
        ' * CHANGE HISTORY:                                                         *
        ' *                                                                         *
        ' * COMPATABILITY: RD-2X/3X RS-1X RS-712                                    *
        ' ***************************************************************************
        Dim RDMessage As String
        RDMessage = New String(" ", RAD_SIZE_MESSAGE)

        'get last error
        Call RadRDStatus(comDevice, Status)
        If Status = 0 Then Exit Sub 'no error found

        Call RadRDMessage(Status, RDMessage)
        Call MsgBox(RDMessage, MsgBoxStyle.Information, "Error")
    End Sub
End Module