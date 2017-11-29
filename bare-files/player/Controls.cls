VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Controls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Exit Sub
    Dim s As String
    s = GetName(Target)
    If s <> vbNullString And s <> "BTNNONE" Then
        Select Case s
            Case "BTNPLAY"
                Call Playlist.StartCurrent
            Case "BTNSTOP"
                Call Playlist.StopPlayback
            Case "BTNPAUSE"
                Call Playlist.PausePlayback
            Case "BTNNEXT"
                Call Playlist.NextItem
            Case "BTNPREV"
                Call Playlist.PrevItem
            Case "S30B"
                Call Player.SeekPlayback(-30)
            Case "S3B"
                Call Player.SeekPlayback(-3)
            Case "S30F"
                Call Player.SeekPlayback(30)
            Case "S3F"
                Call Player.SeekPlayback(3)
        End Select
    End If
    If s <> "BTNNONE" Then
        Call Controls.Range("BTNNONE").Select
    End If
End Sub

'Sets the title
Public Sub SetTitle(ByVal Title As String)
    Call Controls.Unprotect("1234")
    Controls.Range("DISPLAY").Value = Title
    Call Controls.Protect("1234")
End Sub

'Sets the Time
Public Sub SetTime(ByVal pos As Long, ByVal Length As Long)
    If ActiveSheet.Name = Controls.Name Then
        Call Controls.Unprotect("1234")
        If pos > -1 And Length > -1 Then
            Controls.Range("DIAGPOS").Value = CStr(pos)
            Controls.Range("DIAGLEN").Value = CStr(Length)
            Controls.Range("POSITION").Value = Player.GetTime(pos)
            Controls.Range("LENGTH").Value = Player.GetTime(Length)
        Else
            Controls.Range("POSITION").Value = "--:--:--"
            Controls.Range("LENGTH").Value = "--:--:--"
        End If
        Call Controls.Protect("1234")
    End If
End Sub

Private Function GetName(ByVal Target As Range) As String
    GetName = vbNullString
    'Displays
    If Target.Column = 4 And Target.Row = 3 Then GetName = "DISPLAY": Exit Function
    If Target.Column = 12 And Target.Row = 3 Then GetName = "POSITION": Exit Function
    If Target.Column = 13 And Target.Row = 3 Then GetName = "LENGTH": Exit Function
    'Player Row 1
    If Target.Column = 5 And Target.Row = 4 Then GetName = "BTNPLAY": Exit Function
    If Target.Column = 6 And Target.Row = 4 Then GetName = "BTNPAUSE": Exit Function
    If Target.Column = 7 And Target.Row = 4 Then GetName = "BTNSTOP": Exit Function
    'Player row 2
    If Target.Column = 5 And Target.Row = 5 Then GetName = "BTNPREV": Exit Function
    If Target.Column = 6 And Target.Row = 5 Then GetName = "BTNNONE": Exit Function
    If Target.Column = 7 And Target.Row = 5 Then GetName = "BTNNEXT": Exit Function
    'Seek ccontrol
    If Target.Column = 4 And Target.Row = 26 Then GetName = "S30B": Exit Function
    If Target.Column = 5 And Target.Row = 26 Then GetName = "S3B": Exit Function
    If Target.Column = 7 And Target.Row = 26 Then GetName = "S3F": Exit Function
    If Target.Column = 8 And Target.Row = 26 Then GetName = "S30F": Exit Function
End Function


