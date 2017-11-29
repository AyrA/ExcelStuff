VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Playlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const SEL_PLAY As Integer = 1
Const SEL_STOP As Integer = 2
Const SEL_PAUSE As Integer = 3

'Pause status
Private Paused As Boolean
'Track index
Private TI As Integer

Private Property Get TrackIndex() As Integer
    If TI > 0 And TI <= NumTracks Then
        TrackIndex = TI
    Else
        TI = 1
        TrackIndex = 1
    End If
End Property

Private Property Let TrackIndex(ByVal newIndex As Integer)
    If newIndex > 0 And newIndex <= NumTracks Then
        TI = newIndex
    Else
        TI = 1
    End If
End Property

'Selects an entry from the list and colors it
Private Sub SelectEntry(ByVal ID As Integer, ByVal SelectType As Integer)
    Playlist.Cells.Interior.ColorIndex = 0
    
    If SelectType = SEL_PLAY Then Playlist.Cells(ID, 1).Interior.Color = RGB(&HDD, &HFF, &HDD): Exit Sub
    If SelectType = SEL_PAUSE Then Playlist.Cells(ID, 1).Interior.Color = RGB(&HFF, &HFF, &HDD): Exit Sub
    If SelectType = SEL_STOP Then Playlist.Cells(ID, 1).Interior.Color = RGB(&HFF, &HDD, &HDD): Exit Sub
End Sub

'plays thecurrent item
Public Sub StartCurrent()
    StartPlayback (TrackIndex)
End Sub

'Plays a specified Item
Public Sub StartPlayback(ByVal ID As Integer)
    Paused = False
    
    If NumTracks = 0 Then
        MsgBox "Plase add songs to the Playlist first"
        Exit Sub
    End If
    
    If ID < 1 Or ID > NumTracks Then ID = 1
    TrackIndex = ID
    Settings.Range("S_CURRENT").Value = CStr(ID)
    Call SelectEntry(ID, SEL_PLAY)
    Call Player.PlayFile(Playlist.Cells(ID, 1).Value)
    Call Controls.SetTitle(CurrentName)
End Sub

'Pauses Playback
Public Sub PausePlayback()
    If Paused Then
        Call SelectEntry(TrackIndex, SEL_PLAY)
    Else
        Call SelectEntry(TrackIndex, SEL_PAUSE)
    End If
    Paused = Not Paused
    Call Player.PausePlayback(Paused)
End Sub

'Stops playback
Public Sub StopPlayback()
    Paused = False
    Call Player.StopPlayback
    Call SelectEntry(TrackIndex, SEL_STOP)
End Sub

'Initializes variables from settings
Public Sub Init()
    On Error Resume Next
    TrackIndex = 1
    TrackIndex = CInt(Range("S_CURRENT").Value)
End Sub

'Called from Player module or control sheet
Public Sub NextItem()
    TrackIndex = TrackIndex + 1
    If TrackIndex > NumTracks Then TrackIndex = 1
    StartPlayback (TrackIndex)
End Sub

'Called from Player module or control sheet
Public Sub PrevItem()
    TrackIndex = TrackIndex - 1
    If TrackIndex < 1 Then TrackIndex = NumTracks
    StartPlayback (TrackIndex)
End Sub

Public Function NumTracks() As Integer
    Dim I As Integer
    I = 1
    While Playlist.Cells(I, 1) <> ""
        I = I + 1
    Wend
    NumTracks = I - 1
End Function

'Gets the file name without path and extension
Private Function CurrentName() As String
    Dim Current As String
    Current = Playlist.Cells(TrackIndex, 1).Value
    If InStr(1, Current, "\") > 0 And InStrRev(Current, ".") > InStrRev(Current, "\") Then
        Current = Mid(Current, InStrRev(Current, "\") + 1)
        CurrentName = Left(Current, InStrRev(Current, ".") - 1)
    Else
        CurrentName = Current
    End If
End Function
