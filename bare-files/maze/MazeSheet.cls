VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MazeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private x As Integer
Private y As Integer
Private w As Integer
Private h As Integer
Private s As Boolean

Public Sub StartSolve(Width As Integer, Height As Integer)
    x = 2
    y = 2
    w = Width
    h = Height
    Cells(x, y).Select
    s = True
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim offsetX As Integer
    Dim offsetY As Integer
    
    If s Then
        If Target.Cells.count = 1 Then
            If ActiveCell.Interior.ColorIndex = 1 Then
                Cells(x, y).Select
            Else
                If IsAdjacent(ActiveCell.Row, ActiveCell.Column, x, y) Then
                    offsetX = ActiveCell.Row - x
                    offsetY = ActiveCell.Column - y
                    
                    x = ActiveCell.Row
                    y = ActiveCell.Column
                    
                    If Cells(x, y).Interior.ColorIndex < 1 Then
                        Cells(x, y).Interior.ColorIndex = 6
                        If IsCorridor(x, y) Then
                            Cells(x + offsetX, y + offsetY).Select
                        End If
                    ElseIf Cells(x, y).Interior.ColorIndex = 3 Then
                        s = False
                        MsgBox "You win!"
                    Else
                        If IsCorridor(x, y) Then
                            Cells(x + offsetX, y + offsetY).Select
                        End If
                    End If
                Else
                    Cells(x, y).Select
                End If
            End If
        End If
    End If
End Sub

Private Function IsCorridor(x As Integer, y As Integer) As Boolean
    Dim count As Integer
    count = 0
    If Cells(x + 1, y).Interior.ColorIndex = 1 Then
        count = count + 1
    End If
    If Cells(x - 1, y).Interior.ColorIndex = 1 Then
        count = count + 1
    End If
    If Cells(x, y + 1).Interior.ColorIndex = 1 Then
        count = count + 1
    End If
    If Cells(x, y - 1).Interior.ColorIndex = 1 Then
        count = count + 1
    End If
    
    If count > 1 Then
        IsCorridor = True
    Else
        IsCorridor = False
    End If
End Function

Private Function IsAdjacent(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Boolean
    IsAdjacent = False
    If x1 = x2 - 1 Or x1 = x2 + 1 Then
        If y1 = y2 Then
            IsAdjacent = True
        End If
    End If
    If x1 = x2 Then
        If y1 = y2 - 1 Or y1 = y2 + 1 Then
            IsAdjacent = True
        End If
    End If
End Function

Public Sub ClearDoc()
    w = 0
    h = 0
    s = False
    MazeSheet.Cells.Interior.ColorIndex = 0
End Sub

Public Sub ClearSolve()
    Dim i As Integer
    Dim j As Integer
    
    If w > 0 And h > 0 Then
        For i = 1 To w
            For j = 1 To h
                If Cells(i, j).Interior.ColorIndex <> 1 Then
                    Cells(i, j).Interior.ColorIndex = 0
                End If
            Next
        Next
        Cells(2, 2).Interior.ColorIndex = 4
        Cells(h - 1, w - 1).Interior.ColorIndex = 3
        StartSolve w, h
    End If
End Sub
