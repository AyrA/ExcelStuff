VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pX As Integer
Private pY As Integer

Public Property Get x() As Integer
    x = pX
End Property

Public Property Get y() As Integer
    y = pY
End Property

Public Property Let x(newX As Integer)
    pX = newX
End Property

Public Property Let y(newY As Integer)
    pY = newY
End Property
