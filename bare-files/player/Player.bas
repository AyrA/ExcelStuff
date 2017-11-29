Attribute VB_Name = "Player"
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Dim Alias As String

Public Sub StopPlayback()
    Call Send("close all")
    Alias = vbNullString
End Sub

Public Sub PlayFile(ByVal FileName As String)
    Call Send("close all")
    
    'make sure the path is in double quote if it contains spaces
    'You can always supply double quotes, even if not needed if you prefer
    If InStr(FileName, " ") Then
        FileName = Chr(34) & FileName & Chr(34)
    End If
    
    'create unique alias
    Alias = Minute(Now) & Second(Now)
    
    Call Send("Open " & FileName & " ALIAS " & Alias & " Type MPEGVideo WAIT")
    Call Send("Play " & Alias)
    'Start timer
    Call Application.OnTime(Now + TimeValue("00:00:01"), "Player.Refresh")
    
End Sub

Public Sub PausePlayback(ByVal doPause As Boolean)
    If Alias <> vbNullString Then
        If doPause Then
            Call Send("pause " & Alias)
        Else
            Call Send("resume " & Alias)
        End If
    End If
End Sub

Private Function Send(ByVal cmd As String) As String
    Dim ret As String
    ret = String(255, " ")
    Call mciSendString(cmd, ret, 255, 0)
    Send = Trim(ret)
End Function

Public Sub SeekPlayback(ByVal Seconds As Long)
    Dim pos As Long
    If Alias <> vbNullString Then
        pos = Seconds * 1000 + CLng(Send("status " & Alias & " position"))
        If pos >= 0 Then
            Call Send("set " & Alias & " time format milliseconds")
            Call Send("set " & Alias & " seek exactly off")
            Call Send("seek " & Alias & " to " & CStr(pos))
            Call Send("play " & Alias)
        End If
    End If
End Sub


Public Function GetTime(Length As Long) As String
    GetTime = Format(TimeSerial(0, 0, Length \ 1000), "hh:nn:ss")
End Function

Public Sub Refresh()
    Dim L1 As Long
    Dim L2 As Long
    If Alias <> vbNullString Then
        
        L1 = CLng(Send("STATUS " & Alias & " position"))
        L2 = CLng(Send("STATUS " & Alias & " length"))
        Call Controls.SetTime(L1, L2)
        
        If L1 = L2 Then
            Call Playlist.NextItem
        End If
        Call Application.OnTime(Now + TimeValue("00:00:01"), "Player.Refresh")
    Else
        Call Controls.SetTime(-1, -1)
    End If
End Sub
