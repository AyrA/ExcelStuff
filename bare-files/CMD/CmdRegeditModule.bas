Attribute VB_Name = "ModHax"
Option Explicit
'Copyright 2017 /u/AyrA_ch

'Open Excel, Press ALT+F11, add new "module" (not Class module) and paste this entire code in it.

'To use you can either place a button on your excel sheet and hook up these functions to it
'Or you place the cursor inside the Button function you want to use and press F5

'After one successful launch you find the executable on your Desktop.
'You then no longer need this file.

'Launches CMD even if it is disabled
Sub ButtonCmd_Click()
    Dim Dest As String
    Dest = Environ("USERPROFILE") & "\Desktop\cmd.exe"
    
    If DoReplace("DisableCMD", "C:\Windows\System32\cmd.exe", Dest) Then
        Shell Dest, vbNormalFocus
    Else
        MsgBox "The String DisableCMD could not be found", vbExclamation, "Can't replace String"
    End If
End Sub

'Editing Registry Editor no longer seems to work
'(it exits immediately without throwing errors)
Sub ButtonReg_Click()
    Dim Dest As String
    Dest = Environ("USERPROFILE") & "\Desktop\regedit.exe"
    
    If DoReplace("DisableRegistryTools", "C:\Windows\System32\regedit.exe", Dest) Then
        MsgBox "Regedit can't be launched directly because it has a UAC manifest. It's on your desktop"
    Else
        MsgBox "The String DisableRegistryTools could not be found", vbExclamation, "Can't replace String"
    End If
End Sub

'Searches for Unicode Strings in executables and replaces them with a stream of "Q"
Function DoReplace(ByVal Find As String, ByVal SourceFile As String, ByVal DestinationFile As String)
    Dim fileNum As Integer
    Dim bytes() As Byte
    Dim performRepl As Boolean
    Dim i
    Dim j
    DoReplace = False
    
    'Load everything at once
    fileNum = FreeFile
    Open SourceFile For Binary As fileNum
    ReDim bytes(LOF(fileNum) - 1)
    Get fileNum, , bytes
    Close fileNum
    For i = LBound(bytes) To UBound(bytes) - (Len(Find) * 2)
        performRepl = True
        For j = 1 To Len(Find)
            'Check if all bytes match the specified ASCII string
            If bytes(i + (j * 2)) <> Asc(Mid(Find, j, 1)) Then
                performRepl = False
                Exit For
            End If
        Next
        'Replace found string with crap
        If performRepl Then
            For j = 1 To Len(Find)
                bytes(i + (j * 2)) = Asc("Q")
            Next
            DoReplace = True
            Exit For
        End If
    Next
    
    If DoReplace Then
        'Save to new file
        fileNum = FreeFile
        Open DestinationFile For Binary Access Write As fileNum
        Put fileNum, , bytes
        Close fileNum
    End If
End Function

