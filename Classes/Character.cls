VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Character"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''POSITIONS''
Public Property Get MyPosX() As Integer
End Property
Public Property Let MyPosX(v As Integer)
End Property
Public Property Get MyPosY() As Integer
End Property
Public Property Let MyPosY(v As Integer)
End Property
''MODEL''
Public Property Get TheModel() As String
End Property
Public Property Let TheModel(v As String)
End Property
''STATS''
Public Property Get MyAttack() As Integer
End Property
Public Property Let MyAttack(v As Integer)
End Property
Public Property Get Myhp() As Integer
End Property
Public Property Let Myhp(v As Integer)
End Property
 ''MISC''
Public Sub Display()
End Sub
Public Sub DisplayOff()
End Sub
Public Function IsInside(TheRoom As Room) As Boolean
End Function



