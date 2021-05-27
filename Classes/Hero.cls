VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements Character
Option Explicit
Private hItems(0 To 9) As Items '' NEED ADD ARRAY OF ITEM (new class) FOR INVENTORY - array of instance of another class''
Private MyCellX As Integer, MyCellY As Integer 'POSTION
Private hatk As Integer, hdef As Integer, Hhp As Integer 'STATS
Private MyModel As String ' you, :)
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
    Character_MyAttack = hatk + GetModifier()(0)
End Property
Public Property Let Character_MyAttack(v As Integer)
    hatk = v
End Property
Public Property Get Character_Myhp() As Integer
    Character_Myhp = Hhp + GetModifier()(1)
End Property
Public Property Get MyBaseHp() As Integer
    MyBaseHp = Hhp
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

Public Sub AddItemToInventory(item As Items)
Static ItemCount As Integer

If ItemCount > 9 Then 'Call open iventory yadayada you are full blablabalbal
    ReDim hItem(0 To 9)
    ItemCount = 0
End If


Set hItems(ItemCount) = item
ItemCount = ItemCount + 1

End Sub

Private Function GetModifier() As Integer()
Dim i As Integer: i = 0
Dim tmparr(0 To 1) As Integer
On Error GoTo err

For i = 0 To UBound(hItems)
On Error Resume Next
    tmparr(0) = tmparr(0) + hItems(i).Iatk
    tmparr(1) = tmparr(1) + hItems(i).Ihp
Next i


GetModifier = tmparr
Exit Function

err:
GetModifier = Split("0|0", "|")

End Function


