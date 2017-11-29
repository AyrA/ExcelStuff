VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const TOOLBARNAME = "MazeToolBar"

Private Sub Workbook_Open()
    ShowToolbar
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DeleteCommandBar
End Sub

Public Sub ShowToolbar()
' Assumes toolbar not already loaded '
    Application.CommandBars.Add TOOLBARNAME
    AddButton "Create maze", "Generate a Maze", 203, "ThisWorkbook.DoMaze"
    AddButton "Reset solving Process", "Resets the solving process", 128, "ThisWorkbook.ClearSolving"
    AddButton "Clear document", "Clear document", 18, "ThisWorkbook.ClearDocument"
    ' call AddButton more times for more buttons '
    With Application.CommandBars(TOOLBARNAME)
        .Visible = True
        .Position = msoBarTop
    End With
End Sub

Private Sub AddButton(caption As String, tooltip As String, faceId As Long, methodName As String)
Dim Btn As CommandBarButton
    Set Btn = Application.CommandBars(TOOLBARNAME).Controls.Add
    With Btn
        .Style = msoButtonIcon
        .faceId = faceId 'http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
        .OnAction = methodName
        .TooltipText = tooltip
    End With
End Sub

Public Sub DeleteCommandBar()
    Application.CommandBars(TOOLBARNAME).Delete
End Sub

Public Sub ClearDocument()
    MazeSheet.ClearDoc
End Sub

Public Sub ClearSolving()
    MazeSheet.ClearSolve
End Sub

Public Sub DoMaze()
    Dim w As Integer
    Dim h As Integer
    
    w = InputBox("Wie breit?", "Breite", 50)
    h = InputBox("Wie hoch?", "Höhe", 50)
    Call Maze.Generate(w, h)

End Sub
