Attribute VB_Name = "modSpeech"

Public Sub InitializeSpeech()
  
  If MapLoaded = "A1" Then
    Call Character_A1
  End If
  If MapLoaded = "A3" Then
    Call Character_A3
  End If
  If MapLoaded = "A4" Then
    Call Character_A4
  End If
  
End Sub


Public Sub Speech_A1()

  Message(0) = "Sheik's House."
  Message(1) = "Be careful, ill be waiting for you to come back."
  Message(2) = "Hey brother, the White Wizard whants to talk to you, you can find him in the mini island."
  Message(3) = "Yep. I can fix this bridge but i will need some wood. Bring it to me and i will do the rest."
  Message(4) = "You whont be able to leave this island without magic. You should find the black wizards hidden in the island's to learn new spells."
  Message(5) = "You have learned the Cut spell. Now you can use this spell by pressing the Z key. Now head east to start your Quest."
  Message(6) = "Feew!!, this weed is to hard. I wish someone could help me."
  Message(7) = "To the mountin of Wizards."
  Message(8) = "To mini Island."
  Message(9) = "To Centurion Island."
 Message(10) = "Lot of weed grows in the bridge. Only those with special powers will leave this island."
 Message(11) = "You found some wood! This thing doesnt has any use to you but you could try to give it to someone."
 Message(12) = "Good. I see that you got wood, ill fix this bridge in no time."
 Message(13) = "Nothing in there."
 Message(14) = "There should be some wood in one of those jars, go ahead and take a peek."
 Message(15) = "My housband left the island to look for a job, but he hasnt come back. Im really worried about this."
 Message(16) = "You found a Coin! With coins you can buy things to get equiped."
 Message(17) = "Thanks Sheik, i really needed your help. Please recive 10 coins as my gratitud."
 Message(18) = "Thanks Sheik, i really needed your help."
 Message(19) = "Star Gate to Centurion Island."
 
End Sub

Public Sub Character_A1()

  If CharY = 7 Then
    If CharX = 8 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(0)
    End If
  End If
  If CharY = 9 Then
    If CharX = 13 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Mother: " & Message(1)
    End If
  End If
  If CharY = 8 Then
    If CharX = 6 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Brother: " & Message(2)
    End If
  End If
  If CharY = 18 Then
    If CharX = 9 Then
      frmGame.txtSpeech.Visible = True
      If Wood > 0 Then
        frmGame.txtSpeech.Text = "Carpinter: " & Message(12)
        Mid(Map(25), 13, 1) = "G"
        Wood = Wood - 1
      Else
        frmGame.txtSpeech.Text = "Carpinter: " & Message(3)
      End If
    End If
  End If
  If CharY = 19 Then
    If CharX = 30 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "White Wizard: " & Message(4)
    End If
  End If
  If CharY = 35 Then
    If CharX = 31 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Black Wizard: " & Message(5)
      SpellCut = True
    End If
  End If
  If CharY = 10 Then
    If CharX = 32 Then
      If Mid(Map(13), 30, 1) = "G" Then
        If ItemTownFound1 = False Then
          frmGame.txtSpeech.Visible = True
          frmGame.txtSpeech.Text = "Lady: " & Message(17)
          Coin = Coin + 10
          ItemTownFound1 = True
        Else
          frmGame.txtSpeech.Visible = True
          frmGame.txtSpeech.Text = "Lady: " & Message(18)
        End If
      Else
        frmGame.txtSpeech.Visible = True
        frmGame.txtSpeech.Text = "Lady: " & Message(6)
        
      End If
    End If
  End If
  If CharY = 18 Then
    If CharX = 11 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(7)
    End If
  End If
  If CharY = 11 Then
    If CharX = 21 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(8)
    End If
  End If
  If CharY = 8 Then
    If CharX = 38 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(9)
    End If
  End If
  If CharY = 10 Then
    If CharX = 43 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(10)
    End If
  End If
  If CharY = 9 Then
    If CharX = 59 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "The sign reads: " & Message(19)
    End If
  End If
  
End Sub

Public Sub Character_A3()
  If CharY = 10 Then
    If CharX = 9 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Lady: " & Message(14)
    End If
  End If
End Sub

Public Sub Character_A4()
  If CharY = 9 Then
    If CharX = 12 Then
      frmGame.txtSpeech.Visible = True
      frmGame.txtSpeech.Text = "Lady: " & Message(15)
    End If
  End If
End Sub

