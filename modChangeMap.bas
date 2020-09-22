Attribute VB_Name = "modChangeMap"

Public Sub SelectPlace()
  'I really dont have idea of improving this.
  
  'Enter
  If PositionMap = "#" Then
    Call Door1
    Midi = "House.mid"
    Call InitMusic

  End If
  
  'Exit
  If PositionMap = "M" Then
    Call Exit1
    Midi = "Town.mid"
    Call InitMusic
  End If
  
End Sub

Public Sub Door1()

  If MapLoaded = "A1" Then
    If CharY = 7 Then
      If CharX = 10 Then
        Call Map_A2
        MapLoaded = "A2"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
      End If
    End If
    If CharY = 7 Then
      If CharX = 29 Then
        Call Map_A3
        MapLoaded = "A3"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
      End If
    End If
    If CharY = 7 Then
      If CharX = 35 Then
        Call Map_A4
        MapLoaded = "A4"
        CharY = 10
        CharX = 7
        ItemHouseFound1 = False
        ItemHouseFound2 = False
      End If
    End If
  End If
  
End Sub

Public Sub Exit1()

  If MapLoaded = "A2" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 10
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "A3" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 29
        ItemTownFound1 = False
      End If
    End If
  End If

  If MapLoaded = "A4" Then
    If CharY = 10 Then
      If CharX = 7 Then
        Call Map_A1
        MapLoaded = "A1"
        CharY = 7
        CharX = 35
        ItemTownFound1 = False
      End If
    End If
  End If

End Sub
