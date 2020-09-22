Attribute VB_Name = "modItemStuff"

Public Sub TypeOfItem()
  
  If CharFacing = 1 Then
    PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
    If PositionMap = "!" Then Call Jar
  End If
  If CharFacing = 2 Then
    PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
    If PositionMap = "!" Then Call Jar
  End If
  If CharFacing = 3 Then
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    If PositionMap = "!" Then Call Jar
  End If
  If CharFacing = 4 Then
    PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    If PositionMap = "!" Then Call Jar
  End If

End Sub

Public Sub Jar()

  Dim RandomNumber As Integer
  'Find Wood
  If MapLoaded = "A3" Then
    If ItemHouseFound1 = False Then
      Randomize Timer
      RandomNumber = Int(Rnd * 4)
      If RandomNumber = 2 Then
        CharFacing = 5
        Item = 1
        Wood = Wood + 1
        frmGame.txtSpeech.Visible = True
        frmGame.txtSpeech.Text = Message(11)
        ItemHouseFound1 = True
      Else
        frmGame.txtSpeech.Visible = True
        frmGame.txtSpeech.Text = Message(13)
      End If
    Else
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = Message(13)
    End If
    'Find a Coin
    If ItemHouseFound2 = False Then
      Randomize Timer
      RandomNumber = Int(Rnd * 4) + 1
      If RandomNumber = 2 Then
        CharFacing = 5
        Item = 2
        Coin = Coin + 1
        frmGame.txtSpeech.Visible = True
        frmGame.txtSpeech.Text = Message(16)
        ItemHouseFound2 = True
      Else
        frmGame.txtSpeech.Visible = True
        frmGame.txtSpeech.Text = Message(13)
      End If
    Else
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = Message(13)
    End If
    
  End If

End Sub
