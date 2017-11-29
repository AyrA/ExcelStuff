Attribute VB_Name = "Maze"
Option Explicit
Option Base 0

Private z() As Byte
Private tmpCell(3) As Cell
Private Cells As New Collection

Public Sub Generate(ByVal Width As Integer, ByVal Height As Integer)
    Dim neighbors() As Cell
    Dim current As New Cell
    Dim i As Integer
    Dim j As Integer
    
    Set tmpCell(0) = CreateCell(0, 1)
    Set tmpCell(1) = CreateCell(0, -1)
    Set tmpCell(2) = CreateCell(1, 0)
    Set tmpCell(3) = CreateCell(-1, 0)
    
    initStack
    
    current.x = 1
    current.y = 1
    
    Cells.Add current
    
    If Width Mod 2 = 0 Then Width = Width + 1
    If Height Mod 2 = 0 Then Height = Height + 1
    
    ReDim z(Width - 1, Height - 1)
    
    Randomize
    
    For i = 0 To Width - 1
        For j = 0 To Height - 1
            z(i, j) = 1
        Next
    Next i
    
    While Cells.count > 0
        z(current.x, current.y) = 0
        neighbors = GetValidNeighbors(current, Width, Height)
        If neighbors(0) Is Nothing And neighbors(1) Is Nothing And neighbors(2) Is Nothing And neighbors(3) Is Nothing Then
            MarkCell current.x, current.y, False
            Set current = Pop()
        Else
            Push current
            Set current = GetRnd(neighbors)
            MarkCell current.x, current.y, True
        End If
    Wend
    DrawExcel
    MazeSheet.StartSolve Width, Height
End Sub

Private Function GetValidNeighbors(centerTile As Cell, ByVal Width As Integer, ByVal Height As Integer) As Cell()
    Dim count As Integer
    Dim i As Integer
    Dim toCheck As New Cell
    Dim validNeighbors(3) As Cell
    
    For i = 0 To 3
        Set toCheck = New Cell
        toCheck.x = centerTile.x + tmpCell(i).x
        toCheck.y = centerTile.y + tmpCell(i).y
        
        If (toCheck.x Mod 2 = 1 Or toCheck.y Mod 2 = 1) And IsInside(toCheck, Width, Height) Then
            If z(toCheck.x, toCheck.y) = 1 And HasThreeWallsIntact(toCheck, Width, Height) Then
                Set validNeighbors(count) = New Cell
                validNeighbors(count).x = toCheck.x
                validNeighbors(count).y = toCheck.y
                count = count + 1
            End If
        End If
    Next
    
    GetValidNeighbors = validNeighbors
End Function

Private Function HasThreeWallsIntact(toCheck As Cell, ByVal w As Integer, ByVal h As Integer)
    Dim count As Integer
    Dim i As Integer
    Dim neighborToCheck As New Cell
    
    count = 0
    
    For i = 0 To 3
        Set neighborToCheck = New Cell
        neighborToCheck.x = toCheck.x + tmpCell(i).x
        neighborToCheck.y = toCheck.y + tmpCell(i).y
        
        If IsInside(neighborToCheck, w, h) Then
            If z(neighborToCheck.x, neighborToCheck.y) = 1 Then
                count = count + 1
            End If
        End If
    Next
    If count = 3 Then
        HasThreeWallsIntact = True
    Else
        HasThreeWallsIntact = False
    End If
End Function

Private Function IsInside(c As Cell, ByVal w As Integer, ByVal h As Integer) As Boolean
    IsInside = (c.x >= 0 And c.y >= 0 And c.x < w And c.y < h)
End Function

''CELL''
Private Function CreateCell(x As Integer, y As Integer)
    Dim a As New Cell
    a.x = x
    a.y = y
    Set CreateCell = a
End Function

''STACK''
Private Sub initStack()
    Set Cells = New Collection
End Sub

Private Sub Push(newItem As Variant)
    With Cells
        .Add newItem
    End With
End Sub

Private Function Pop() As Variant
    With Cells
        If .count > 0 Then
            Set Pop = .Item(.count)
            .Remove .count
        End If
    End With
End Function

''MISC''
Private Function GetRnd(arr)
    Dim a
    Set a = Nothing
    
    While a Is Nothing
        Set a = arr(Rnd * UBound(arr))
    Wend
    
    Set GetRnd = a
End Function

Private Sub DrawExcel()
    Dim ws As Worksheet
    Dim i As Integer
    Dim j As Integer
    
    MazeSheet.Cells.Interior.ColorIndex = 0

    For i = 0 To UBound(z, 1)
        For j = 0 To UBound(z, 2)
            MazeSheet.Cells(i + 1, j + 1).Interior.ColorIndex = z(i, j)
        Next
        DoEvents
    Next
    MazeSheet.Cells(2, 2).Interior.ColorIndex = 4 'Green
    MazeSheet.Cells(i - 1, j - 1).Interior.ColorIndex = 3 'Red
End Sub

Private Sub MarkCell(x As Integer, y As Integer, z As Boolean)
    If z Then
        MazeSheet.Cells(x + 1, y + 1).Interior.ColorIndex = 3 'Red
    Else
        MazeSheet.Cells(x + 1, y + 1).Interior.ColorIndex = 4 'Green
    End If
End Sub

Public Sub Test()
    Call Generate(51, 51)
End Sub

