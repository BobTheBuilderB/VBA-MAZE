Attribute VB_Name = "Maze"
Option Explicit
Option Base 0

Const LabSize As Integer = 100

Private Rooms As New Collection  ''STACK''

Private TmpRooms(3) As Room

Sub Generate()

Application.ScreenUpdating = False
randomize
Sheets(1).UsedRange.ClearContents
Sheets(1).UsedRange.Interior.ColorIndex = 0

Dim Branch As Integer
Branch = Int(LabSize / 100) ''NUMBER OF BACK BEFORE BOSSROM''

''SETUP START ROOM''
Dim i As Integer, counter As Integer: counter = 0 ''counter needs to be relative to maze size to find good lastroom of maze. 1 is nice for 100*100 more than 1 can cause bug
Dim ValidAttempts() As Room
Dim current As New Room

If GameEvents.BossRoom Is Nothing Then
    Call SetRoom(current)
Else
    Set current = GameEvents.BossRoom
End If

Dim TheRoomID As Integer ''BETTER NAME IS DOORCOUNT''
Dim TheRoomIDForEnemies As Integer ''BETTERNAME IS ROOMCOUNT''
TheRoomID = 1
TheRoomIDForEnemies = 1

Dim Fail As Integer ''For door coordination - do not check avaible room on failure first time

Dim PreviousRoom As Room

''FIRST ROOM SETUP''
Call SetTmpRooms(current)
current.SetMyID = TheRoomID
Set PreviousRoom = current
Rooms.Add current
CurrentLevelRooms.Add current
current.Draw
Call current.colored(4)

''CORE LOOP STARTS''
Do While Rooms.count > 0
If Fail <> 1 Then
ValidAttempts = GetValidAttempts(current)
End If

    If ValidAttempts(0) Is Nothing And ValidAttempts(1) Is Nothing And ValidAttempts(2) Is Nothing And ValidAttempts(3) Is Nothing Then

        If counter = Branch Then
            'call clearroom
            Call current.colored(3)
            Set GameEvents.BossRoom = current
'            GameEvents.PopulateBossRoom
        End If
        
        Set current = Stuff.Pop(Rooms)
        counter = counter + 1
        Fail = Fail + 1
        
    Else
        Fail = 0
        
        Set current = Stuff.GetRnd(ValidAttempts)
        current.Draw
                       
        TheRoomID = TheRoomID + 1 ''ROOM WHO BUILD DOOR''
        current.SetMyID = TheRoomID
        Rooms.Add current
        CurrentLevelRooms.Add current
        
        Call WorldGen.PopulateRoom(TheRoomID)
        Call WorldGen.AddDoorToRoom(TheRoomID, PreviousRoom.GetMyID)
    End If
Set PreviousRoom = current
Call SetTmpRooms(current)
Loop

Set current = Nothing
Application.ScreenUpdating = True
End Sub
Private Function GetValidAttempts(CurrentRoom As Room) As Room()
    Dim i As Integer, count As Integer
    Dim ToTry As New Room
    Dim ValidAttempts(3) As Room

count = 0

For i = 0 To 3
    Set ToTry = New Room
    Call AddRooms(ToTry, CurrentRoom)
    Call AddRooms(ToTry, TmpRooms(i))
    Call ToTry.AmIValid(LabSize)
    If ToTry.ValidOrNot = 1 Then
        Set ValidAttempts(count) = New Room
        Call AddRooms(ValidAttempts(count), ToTry)
        count = count + 1
    End If
Next i

GetValidAttempts = ValidAttempts
End Function
''ROOMS''
Private Function SetRoom(theroomyouwanttosett As Room, Optional x1 As Integer = 1, Optional y1 As Integer = 1, Optional x2 As Integer = 1, _
Optional y2 As Integer = 21, Optional x3 As Integer = 21, Optional y3 As Integer = 1, Optional x4 As Integer = 21, Optional y4 As Integer = 21, Optional Directio As Integer = 0)
theroomyouwanttosett.x1 = x1
theroomyouwanttosett.y1 = y1
theroomyouwanttosett.x2 = x2
theroomyouwanttosett.y2 = y2
theroomyouwanttosett.x3 = x3
theroomyouwanttosett.y3 = y3
theroomyouwanttosett.x4 = x4
theroomyouwanttosett.y4 = y4
theroomyouwanttosett.MyDirection = Directio
End Function
Private Function AddRooms(theroomyouwanttoset As Room, addroom As Room)
theroomyouwanttoset.x1 = theroomyouwanttoset.x1 + addroom.x1
theroomyouwanttoset.y1 = theroomyouwanttoset.y1 + addroom.y1
theroomyouwanttoset.x2 = theroomyouwanttoset.x2 + addroom.x2
theroomyouwanttoset.y2 = theroomyouwanttoset.y2 + addroom.y2
theroomyouwanttoset.x3 = theroomyouwanttoset.x3 + addroom.x3
theroomyouwanttoset.y3 = theroomyouwanttoset.y3 + addroom.y3
theroomyouwanttoset.x4 = theroomyouwanttoset.x4 + addroom.x4
theroomyouwanttoset.y4 = theroomyouwanttoset.y4 + addroom.y4
theroomyouwanttoset.MyDirection = addroom.MyDirection
End Function
''RANDOMIZE ROOMS''
Private Sub SetTmpRooms(CurrentRoom As Room)
Dim Longeur As Integer, largeur As Integer

If Not IsEmpty(TmpRooms) Then Erase TmpRooms

Longeur = Int(20 * rnd + 7) 'Columns
largeur = Int(20 * rnd + 7)  ' rows
'up
Set TmpRooms(0) = New Room
Call SetRoom(TmpRooms(0), -largeur, 0, -largeur, 0, CurrentRoom.x1 - CurrentRoom.x3 - 1, 0, CurrentRoom.x1 - CurrentRoom.x3 - 1, 0, 0)
'right
Set TmpRooms(1) = New Room
Call SetRoom(TmpRooms(1), 0, CurrentRoom.y2 - CurrentRoom.y1 + 1, 0, Longeur, 0, CurrentRoom.y2 - CurrentRoom.y1 + 1, 0, Longeur, 1)
'bot
Set TmpRooms(2) = New Room
Call SetRoom(TmpRooms(2), CurrentRoom.x3 - CurrentRoom.x1 + 1, 0, CurrentRoom.x3 - CurrentRoom.x1 + 1, 0, largeur, 0, largeur, 0, 2)
'left
Set TmpRooms(3) = New Room
Call SetRoom(TmpRooms(3), 0, -Longeur, 0, CurrentRoom.y1 - CurrentRoom.y2 - 1, 0, -Longeur, 0, CurrentRoom.y1 - CurrentRoom.y2 - 1, 3)
End Sub

