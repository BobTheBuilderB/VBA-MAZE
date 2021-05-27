VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Door"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cell1X As Integer, cell1Y As Integer, cell2X As Integer, cell2Y As Integer ''POSITIONS''
Private IsOpen As Boolean ''OPEN OR CLOSED''
''POSITIONS''
Public Property Get x1() As Integer
    x1 = cell1X
End Property
Public Property Get y1() As Integer
    y1 = cell1Y
End Property
Public Property Get x2() As Integer
    x2 = cell2X
End Property
Public Property Get y2() As Integer
    y2 = cell2Y
End Property
Public Property Let x1(newx1 As Integer)
   cell1X = newx1
End Property
Public Property Let y1(newy1 As Integer)
    cell1Y = newy1
End Property
Public Property Let x2(newx2 As Integer)
    cell2X = newx2
End Property
Public Property Let y2(newy2 As Integer)
    cell2Y = newy2
End Property

Sub DisplayOpen()
Cells(cell1X, cell1Y).Value = ""
Cells(cell2X, cell2Y).Value = ""
End Sub
Sub DisplayClosed()
Cells(cell1X, cell1Y).Value = "D"
Cells(cell2X, cell2Y).Value = "D"
End Sub
