Attribute VB_Name = "GameCalcs"
Option Explicit

''FIND ENEMY INDEX FROM ENEMIES COLLECTION''
Public Function FindEnemyIndex(X As Integer, Y As Integer, TheRoomYouAreIn As Room) As Integer
Dim lpchar As Character

Dim count As Integer
count = 1 ''COLLECTIONS INITIALIZE AT 1''

    For Each lpchar In TheRoomYouAreIn.GetEnemies
        If lpchar.MyPosX = X And lpchar.MyPosY = Y Then
            FindEnemyIndex = count
            Exit For
        Else
        count = count + 1
        End If
    Next lpchar
End Function

