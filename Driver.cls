VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
''START GAME''

Sub initialize()
'UserForm1.Show vbModeless
Call GameEvents.init
End Sub

Private Sub Workbook_Activate()
 Application.OnKey "{LEFT}", "ThisWorkbook.OnLeftArrowKeyPress"
 Application.OnKey "{RIGHT}", "ThisWorkbook.OnRightArrowKeyPress"
 Application.OnKey "{UP}", "ThisWorkbook.OnUpArrowKeyPress"
 Application.OnKey "{DOWN}", "ThisWorkbook.OnDownArrowKeyPress"
End Sub


Private Sub Workbook_Deactivate()
 Application.OnKey "{LEFT}", ""
 Application.OnKey "{RIGHT}", ""
 Application.OnKey "{UP}", ""
 Application.OnKey "{DOWN}", ""
End Sub
'Inputs
Sub OnLeftArrowKeyPress()
Call GameEvents.Move(Hero, 0, -1)
If Hero.Character_IsInside(GameEvents.BossRoom) = True Then
    MsgBox "You Won"
    Call Maze.Generate
End If
End Sub
Sub OnRightArrowKeyPress()
Call GameEvents.Move(Hero, 0, 1)
If Hero.Character_IsInside(GameEvents.BossRoom) = True Then
    MsgBox "You Won"
    Call Maze.Generate
End If
End Sub
Sub OnUpArrowKeyPress()
Call GameEvents.Move(Hero, -1, 0)
If Hero.Character_IsInside(GameEvents.BossRoom) = True Then
    MsgBox "You Won"
    Call Maze.Generate
End If
End Sub
Sub OnDownArrowKeyPress()
Call GameEvents.Move(Hero, 1, 0)
If Hero.Character_IsInside(GameEvents.BossRoom) = True Then
    MsgBox "You Won"
    Call Maze.Generate
End If
End Sub


