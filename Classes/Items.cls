VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim atk As Integer, HP As Integer
Dim ID As Integer
Dim Name As String
Public Property Get Iatk()
Iatk = atk
End Property
Public Property Get Ihp()
Ihp = HP
End Property
Public Property Get Iname()
Iname = Name
End Property
Public Property Get Iid()
Iid = ID
End Property

Public Sub rnd()
atk = 5
HP = 5
Name = Stuff.RandomNameGenerator
End Sub
