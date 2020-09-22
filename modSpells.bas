Attribute VB_Name = "modSpells"

Public Sub Make_Spell_Cut()
  
  If CharFacing = 1 Then 'Up
    PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
    If PositionMap = "<" Then
      Mid(Map(CharY + 2), CharX + 3, 1) = "G"
    End If
  End If
  
  If CharFacing = 2 Then 'Down
    PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
    If PositionMap = "<" Then
      Mid(Map(CharY + 4), CharX + 3, 1) = "G"
    End If
  End If
  
  If CharFacing = 3 Then 'Right
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    If PositionMap = "<" Then
      Mid(Map(CharY + 3), CharX + 4, 1) = "G"
    End If
  End If
  
  If CharFacing = 4 Then 'Left
    PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    If PositionMap = "<" Then
      Mid(Map(CharY + 3), CharX + 2, 1) = "G"
    End If
  End If
End Sub
