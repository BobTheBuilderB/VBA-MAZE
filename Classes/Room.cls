VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Enemies As New Collection
Private Doors As New Collection

Private TLr As Integer, TLc As Integer, TRr As Integer, TRc As Integer, BLr As Integer, BLc As Integer, BRr As Integer, BRc As Integer 'position

Private ID As Integer
Private IsValid As Integer
Private direction As Integer
''ID''
Public Property Get GetMyID() As Integer
    GetMyID = ID
End Property
Public Property Let SetMyID(v As Integer)
    ID = v
End Property
''DIRECTION''
Public Property Get MyDirection() As Integer
    MyDirection = direction
End Property
''DIRECTION''
Public Property Get GetDoors() As Collection
    Set GetDoors = Doors
End Property
Public Property Let MyDirection(MyDirection2 As Integer)
    direction = MyDirection2
End Property
''VALIDITY''
Public Property Get ValidOrNot() As Integer
    ValidOrNot = IsValid
End Property
''AVAIBLEAREA''
Public Property Get GetAvaibleArea() As Range
    Set GetAvaibleArea = AvaibleArea
End Property
''ENEMIES''
Public Property Get GetEnemies() As Collection
    Set GetEnemies = Enemies
End Property
''POSITIONS''
Public Property Get x1() As Integer
    x1 = TLr
End Property
Public Property Get y1() As Integer
    y1 = TLc
End Property
Public Property Get x2() As Integer
    x2 = TRr
End Property
Public Property Get y2() As Integer
    y2 = TRc
End Property
Public Property Get x3() As Integer
    x3 = BLr
End Property
Public Property Get y3() As Integer
    y3 = BLc
End Property
Public Property Get x4() As Integer
    x4 = BRr
End Property
Public Property Get y4() As Integer
    y4 = BRc
End Property
Public Property Let x1(newx1 As Integer)
   TLr = newx1
End Property
Public Property Let y1(newy1 As Integer)
    TLc = newy1
End Property
Public Property Let x2(newx2 As Integer)
    TRr = newx2
End Property
Public Property Let y2(newy2 As Integer)
    TRc = newy2
End Property
Public Property Let x3(newx3 As Integer)
    BLr = newx3
End Property
Public Property Let y3(newy3 As Integer)
    BLc = newy3
End Property
Public Property Let x4(newx4 As Integer)
    BRr = newx4
End Property
Public Property Let y4(newy4 As Integer)
    BRc = newy4
End Property
Public Sub Draw()
Dim i As Integer, j As Integer, Longeur As Integer, largeur As Integer 'longeur = x largeur = y
i = 0: j = 0
Longeur = TRc - TLc: largeur = BLr - TLr ''PRIVATE FUNCTION CALCRANGE
    For i = 0 To Longeur
        Cells(TLr, TLc + i).Value = "W"
        Cells(BLr, BLc + i).Value = "W"
    Next i
    
    For j = 1 To largeur - 1
        Cells(TLr + j, TLc).Value = "W"
        Cells(TRr + j, TRc).Value = "W"
    Next j
    
'Dim dr As Door
'    For Each dr In Doors
'       dr.DisplayClosed
'    Next dr
          
End Sub
Public Sub Flavor()
'FLAVOR THE ROOM
End Sub
Public Sub colored(color As Integer)
Dim Longeur As Integer, largeur As Integer, i As Byte, j As Byte
Longeur = TRc - TLc: largeur = BLr - TLr
i = 0: j = 0
        For i = 1 To Longeur - 1
            For j = 1 To largeur - 1
                Cells(TLr + j, TLc + i).Interior.ColorIndex = color
            Next j
        Next i
End Sub
Public Sub StoreDoor(TheDoor As Door)
Doors.Add TheDoor
End Sub
Public Function GetDoor(i As Integer) As Door
Set GetDoor = Doors(i)
End Function
Public Sub StoreEnemy(TheEnemy As Enemy)
Enemies.Add TheEnemy
End Sub
Public Sub RemoveEnemy(index As Integer)
Enemies.Remove index
End Sub
Public Function AmIValid(LabSize As Integer) As Integer
Dim i As Integer, j As Integer, Longeur As Integer, largeur As Integer 'longeur = x largeur = y
i = 0: j = 0
IsValid = 1
Longeur = TRc - TLc: largeur = BLr - TLr
If Longeur < 3 Or largeur < 3 Or TRc < 1 Or TLc < 1 Or BLr < 1 Or TLr < 1 Or BLr > 1 + LabSize Or BRc > 1 + LabSize Then
    IsValid = 2
    Else
        For i = 0 To Longeur
            If Cells(TLr, TLc + i).Value = "W" Or Cells(BLr, BLc + i).Value = "W" Then IsValid = 2
        Next i
        For j = 1 To largeur - 1
            If Cells(TLr + j, TLc).Value = "W" Or Cells(TRr + j, TRc).Value = "W" Then IsValid = 2
        Next j
    End If
AmIValid = IsValid
End Function
Private Function AvaibleArea() As Range 'USABLE AREA - AVAIBLE AREA IS -1 +1, ROOM INCLUDED WALLS AREA IS -0 +0
Set AvaibleArea = Range(Cells(TLr + 2, TLc + 2).Address, Cells(BRr - 2, BRc - 2).Address)
End Function

