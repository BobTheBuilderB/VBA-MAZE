VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fight 
   Caption         =   "UserForm1"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   OleObjectBlob   =   "Fight.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Fight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim eIndex As Integer
Dim FightRoom As Room


Public Sub LoadChars(h As Character, i As Integer, r As Room)
eIndex = i

Set FightRoom = r

Me.HeroLabel.Caption = h.TheModel
Me.EnemyLabel.Caption = FightRoom.GetEnemies(eIndex).Character_TheModel

Me.ATK1.Caption = h.MyAttack
Me.ATK2.Caption = FightRoom.GetEnemies(eIndex).Character_MyAttack

Me.HP1.Caption = h.Myhp
Me.HP2.Caption = FightRoom.GetEnemies(eIndex).Character_Myhp
End Sub

Private Sub CommandButton1_Click()

FightRoom.GetEnemies(eIndex).Character_Myhp = FightRoom.GetEnemies(eIndex).Character_Myhp - Hero.Character_MyAttack
Hero.Character_Myhp = Hero.MyBaseHp - FightRoom.GetEnemies(eIndex).Character_MyAttack

Me.HP1.Caption = Hero.Character_Myhp
Me.HP2.Caption = FightRoom.GetEnemies(eIndex).Character_Myhp

If Hero.Character_Myhp <= 0 Then
    MsgBox ("you lose")
    Call GameEvents.init
End If

If FightRoom.GetEnemies(eIndex).Character_Myhp <= 0 Then
    Call Hero.Character_Display
    FightRoom.GetEnemies(eIndex).Character_DisplayOff
    Call FightRoom.RemoveEnemy(eIndex)

    Dim i As Integer: i = 1
    For i = 1 To FightRoom.GetDoors.count
    FightRoom.GetDoor(i).DisplayOpen
    Next i
    
    Debug.Print FightRoom.GetDoors.count
    
    Set FightRoom = Nothing
    Unload Me
End If
End Sub

Private Sub UserForm_Click()

End Sub
