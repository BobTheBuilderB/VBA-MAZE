Attribute VB_Name = "Stuff"
        ''Stuff for second axis randomization on maze''
        'If randomnumber > 0.49 Then
        '    Call SetRoom(TmpRooms(i), -largeur, CurrentRoom.y2 - CurrentRoom.y1 + 1, -largeur, longeur, largeur, CurrentRoom.y2 - CurrentRoom.y1 + 1, largeur, longeur)
        'Else
        '    If largeur > CurrentRoom.x3 - CurrentRoom.x1 Then largeur = 1
        '    Call SetRoom(TmpRooms(i), largeur, CurrentRoom.y2 - CurrentRoom.y1 + 1, largeur, longeur, -largeur, CurrentRoom.y2 - CurrentRoom.y1 + 1, -largeur, longeur)
        'End If
        'If randomnumber > 0.49 Then
        '    Call SetRoom(TmpRooms(i), -largeur, -longeur, -largeur, longeur, CurrentRoom.x1 - CurrentRoom.x3 - 1, -longeur, CurrentRoom.x1 - CurrentRoom.x3 - 1, longeur)
        'Else
        '    If longeur > CurrentRoom.y2 - CurrentRoom.y1 Then longeur = 1
        '    Call SetRoom(TmpRooms(i), -largeur, longeur, -largeur, longeur, CurrentRoom.x1 - CurrentRoom.x3 - 1, longeur, CurrentRoom.x1 - CurrentRoom.x3 - 1, -longeur)
        'End If
        
           
''EMPTY COLLECTION''
Sub limpieza(ByRef listilla As Collection)
    While listilla.count <> 0
        listilla.Remove (listilla.count)
    Wend
End Sub
''RANDOM FROM ARRAY''
Function GetRnd(arr)
    Dim a
    Set a = Nothing
    
    While a Is Nothing
        Set a = arr(rnd * UBound(arr))
    Wend
    
    Set GetRnd = a
End Function
''STACK''
Function Pop(WhichCollection As Collection) As Variant
    With WhichCollection
        If .count > 0 Then
            Set Pop = .item(.count)
            .Remove .count
        End If
    End With
End Function
Function AreWeLooping(r As Room, c As Collection) As Boolean
If r Is c(c.count) Then
    AreWeLooping = True
Else
    AreWeLooping = False
End If
End Function

Function RandomNameGenerator() As String
'PURPOSE: Create a Randomized String of Characters
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim RND1 As Variant
Dim RND2 As Variant

  RND1 = Array("Sword", "Axe", "Spear", "Dildo", "Pickaxe", "Stick", "Mace", "Hammer", "Gun", "Sniper", _
  "Sextoy", "Strap-On", "Trebuchet", "Catapult", "Balista")

  RND2 = Array("Destruction", "Thousand thuth", "Pleasure", "Joy", "Mightiness", "Reckoning", "Fear", "Pain", "Wood", "Iron", _
  "Mining", "Agility", "Precision", "Superiority", "Defeat", "Death")
  

randomize
RandomNameGenerator = RND1(Int((UBound(RND1) - LBound(RND1) + 1) * rnd + LBound(RND1))) & " " & "Of" & " " & RND2(Int((UBound(RND2) - LBound(RND2) + 1) * rnd + LBound(RND2)))

End Function
