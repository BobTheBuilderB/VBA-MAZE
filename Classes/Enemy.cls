VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Character
Option Explicit

Private MyCellX As Integer, MyCellY As Integer 'POSTION
Private hatk As Integer, Hhp As Integer 'STATS
Private MyModel As String


''''' CHARACTER INTERFACE ''''
''POSITIONS''
Public Property Get Character_MyPosX() As Integer
    Character_MyPosX = MyCellX
End Property
Public Property Let Character_MyPosX(v As Integer)
    MyCellX = v
End Property
Public Property Get Character_MyPosY() As Integer
    Character_MyPosY = MyCellY
End Property
Public Property Let Character_MyPosY(v As Integer)
    MyCellY = v
End Property
''MODEL''
Public Property Get Character_TheModel() As String
    Character_TheModel = MyModel
End Property
Public Property Let Character_TheModel(v As String)
    MyModel = v
End Property
''STATS''
Public Property Get Character_MyAttack() As Integer
    Character_MyAttack = hatk
End Property
Public Property Let Character_MyAttack(v As Integer)
    hatk = v
End Property
Public Property Get Character_Myhp() As Integer
    Character_Myhp = Hhp
End Property
Public Property Let Character_Myhp(v As Integer) 'For dmg
    Hhp = v
End Property

Public Sub Character_Display()
Cells(MyCellX, MyCellY).Value = Me.Character_TheModel
End Sub

Public Sub Character_DisplayOff()
Cells(MyCellX, MyCellY).Value = ""
End Sub

Public Function Character_IsInside(TheRoom As Room) As Boolean

If Me.Character_MyPosY > TheRoom.y1 And Me.Character_MyPosY < TheRoom.y2 And Me.Character_MyPosX > TheRoom.x1 And Me.Character_MyPosX < TheRoom.x3 Then
    Character_IsInside = True
    Else
    Character_IsInside = False
End If

End Function







