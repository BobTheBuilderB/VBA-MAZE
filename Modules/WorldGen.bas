Attribute VB_Name = "WorldGen"
Option Explicit

Public Sub AddDoorToRoom(roomwhobuilddoor As Integer, roomwhoowndoor As Integer)
    Dim CurrentDoor As New Door
    Call SetDoor(CurrentDoor, CurrentLevelRooms(roomwhobuilddoor))
    
    Call CurrentLevelRooms(roomwhoowndoor).StoreDoor(CurrentDoor)

    If roomwhobuilddoor - 1 = 1 Then
        Call CurrentDoor.DisplayOpen
    Else
        Call CurrentDoor.DisplayClosed
    End If
End Sub
Private Sub SetDoor(TheDoor As Door, TheRoom As Room) ''NEED IMPROVEMENT''
Dim Longeur As Integer, largeur As Integer 'longeur = x largeur = y
Longeur = TheRoom.y2 - TheRoom.y1: largeur = TheRoom.x3 - TheRoom.x1

Select Case TheRoom.MyDirection
    Case 0 'UP
        TheDoor.x1 = TheRoom.x3
        TheDoor.y1 = TheRoom.y1 + Int(Longeur / 2)
        TheDoor.x2 = TheRoom.x3 + 1
        TheDoor.y2 = TheRoom.y1 + Int(Longeur / 2)

    Case 1 'RIGHT
        TheDoor.x1 = TheRoom.x1 + Int(largeur / 2)
        TheDoor.y1 = TheRoom.y3
        TheDoor.x2 = TheRoom.x1 + Int(largeur / 2)
        TheDoor.y2 = TheRoom.y3 - 1

    Case 2 'BOTTOM
        TheDoor.x1 = TheRoom.x1
        TheDoor.y1 = TheRoom.y1 + Int(Longeur / 2)
        TheDoor.x2 = TheRoom.x1 - 1
        TheDoor.y2 = TheRoom.y1 + Int(Longeur / 2)

    Case 3 'LEFT
        TheDoor.x1 = TheRoom.x1 + Int(largeur / 2)
        TheDoor.y1 = TheRoom.y2
        TheDoor.x2 = TheRoom.x1 + Int(largeur / 2)
        TheDoor.y2 = TheRoom.y2 + 1
End Select
End Sub

Public Sub PopulateRoom(RoomID As Integer)
    Dim Longeur As Integer, largeur As Integer
    
    Longeur = CurrentLevelRooms(RoomID).y4 - CurrentLevelRooms(RoomID).y3
    largeur = CurrentLevelRooms(RoomID).x3 - CurrentLevelRooms(RoomID).x1
    
    Do
    Dim CurrentEnemy As New Enemy
    Call GameEvents.SetChar(CurrentEnemy, CurrentLevelRooms(RoomID).x1 + Int((largeur - 2) * rnd + 2), CurrentLevelRooms(RoomID).y1 + Int((Longeur - 2) * rnd + 2), "E")
    Loop While Intersect(CurrentLevelRooms(RoomID).GetAvaibleArea, Cells(CurrentEnemy.Character_MyPosX, CurrentEnemy.Character_MyPosY)) Is Nothing
    
    Call CurrentLevelRooms(RoomID).StoreEnemy(CurrentEnemy)
    Call CurrentEnemy.Character_Display
    
End Sub
Public Sub PopulateBossRoom()

End Sub


