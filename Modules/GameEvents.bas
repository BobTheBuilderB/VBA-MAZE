Attribute VB_Name = "GameEvents"
Option Explicit

''MAZEROOMS''
Public CurrentLevelRooms As Collection
''BOSSROOM''
Public BossRoom As Room

''LOOPS''
Private lpchar As Character
Private lpdoor As Door

''INITIALIZE GAME''
Public Sub init()

    If Not CurrentLevelRooms Is Nothing Then
        Call limpieza(CurrentLevelRooms)
    End If
    
    Set CurrentLevelRooms = New Collection
    Set BossRoom = Nothing

    Call Maze.Generate
    Call SetChar(Hero, 10, 10, "@")
    
    Dim testitem As New Items
    Call testitem.rnd
    Call Hero.AddItemToInventory(testitem)
    
        Dim a As New Items
    Call a.rnd
    Call Hero.AddItemToInventory(a)
    
        Dim b As New Items
    Call b.rnd
    Call Hero.AddItemToInventory(b)
    
        Dim c As New Items
    Call c.rnd
    Call Hero.AddItemToInventory(c)
End Sub
Public Sub Move(Who As Character, X As Integer, Y As Integer)
Dim enemyindex As Integer

Call Who.DisplayOff

Select Case Cells(Who.MyPosX + X, Who.MyPosY + Y).Value
    Case "W", "D"
        Who.Display
        
    Case "E", "S", "G"
         enemyindex = GameCalcs.FindEnemyIndex(Who.MyPosX + X, Who.MyPosY + Y, CurrentLevelRooms(WhichRoomAmIIn(Hero)))
         Call OpenFightMenu(Who, enemyindex, CurrentLevelRooms(WhichRoomAmIIn(Hero)))

    Case Else
        Who.MyPosX = Who.MyPosX + X
        Who.MyPosY = Who.MyPosY + Y
        Call Who.Display
                    
End Select
End Sub
Public Sub RemoveEnemyFromCollection(index As Integer)
    Call Enemies(index).Character_DisplayOff
    Call Enemies.Remove(index)
    Doors(index + 1).DisplayOpen
End Sub
Public Sub SetChar(char As Character, X As Integer, Y As Integer, Optional Model As String)
    char.MyPosX = X
    char.MyPosY = Y
    char.TheModel = Model
    
    Select Case Model
        Case "E"
            char.MyAttack = 1
            char.Myhp = 3
        Case "B"
            char.MyAttack = 2
            char.Myhp = 7
    End Select
End Sub
Private Sub OpenFightMenu(h As Character, enemyindex As Integer, r As Room)
    Call Fight.LoadChars(h, enemyindex, r)
    Fight.Show vbModeless
End Sub

Private Sub updatos()

End Sub

Private Function WhichRoomAmIIn(c As Character) As Integer
Dim i As Integer
i = 1

For i = 1 To CurrentLevelRooms.count
    If c.IsInside(CurrentLevelRooms(i)) = True Then
        WhichRoomAmIIn = i
        Exit Function
    End If
Next i
End Function

