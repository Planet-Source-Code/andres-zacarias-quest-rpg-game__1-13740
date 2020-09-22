VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quest"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5760
      Top             =   4080
   End
   Begin VB.TextBox txtSpeech 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmGame.frx":0000
      Top             =   3720
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5880
      Top             =   3240
   End
   Begin VB.Image imgItem2 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":000A
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarryChar 
      Height          =   480
      Left            =   6120
      Picture         =   "frmGame.frx":0C4C
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgItem1 
      Height          =   480
      Left            =   6360
      Picture         =   "frmGame.frx":188E
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen3 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":24D0
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen2 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":3112
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "X"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgBed2 
      Height          =   480
      Left            =   5400
      Picture         =   "frmGame.frx":3D54
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBed1 
      Height          =   480
      Left            =   5400
      Picture         =   "frmGame.frx":4996
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTable2 
      Height          =   480
      Left            =   4920
      Picture         =   "frmGame.frx":55D8
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTable1 
      Height          =   480
      Left            =   4440
      Picture         =   "frmGame.frx":621A
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExitB2 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":6E5C
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExitB1 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":7A9E
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgJar 
      Height          =   480
      Left            =   4440
      Picture         =   "frmGame.frx":86E0
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouseExit 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":9322
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":9F64
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL3 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":ABA6
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL3 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":B7E8
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT3 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":C42A
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":D06C
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR3 
      Height          =   495
      Left            =   3000
      Picture         =   "frmGame.frx":DCAE
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR3 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":E950
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB3 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":F592
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWall2 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":101D4
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeed 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":10E16
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBadWizard 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":11A58
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharGoodWizard 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":1269A
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy2 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":132DC
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":13F1E
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":14B60
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL2 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":157A2
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR2 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":163E4
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOR2 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":17026
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOL2 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":17C68
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFloor1 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":188AA
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallTop 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":194EC
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallBottom 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":1A12E
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":1AD70
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":1B9B2
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":1C5F4
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT2 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":1D236
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":1DE78
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmGame.frx":1EABA
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmGame.frx":1F6FC
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB2 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":2033E
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRedTop 
      Height          =   480
      Left            =   5040
      Picture         =   "frmGame.frx":20F80
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDoor1 
      Height          =   480
      Left            =   5040
      Picture         =   "frmGame.frx":21BC2
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgNothing 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":22804
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   320
      X2              =   320
      Y1              =   320
      Y2              =   0
   End
   Begin VB.Image imgBlueTop 
      Height          =   480
      Left            =   4560
      Picture         =   "frmGame.frx":23446
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWindow1 
      Height          =   480
      Left            =   4560
      Picture         =   "frmGame.frx":24088
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharWomen1 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":24CCA
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharBoy1 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":2590C
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDownChar 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":2654E
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRightChar 
      Height          =   480
      Left            =   5640
      Picture         =   "frmGame.frx":27190
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLeftChar 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":27DD2
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgUpChar 
      Height          =   480
      Left            =   5160
      Picture         =   "frmGame.frx":28A14
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIB 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":29656
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOL 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":2A298
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBOR 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":2AEDA
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIR 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":2BB1C
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIR 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":2C75E
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIL 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":2D3A0
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIT 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":2DFE2
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOR 
      Height          =   480
      Left            =   600
      Picture         =   "frmGame.frx":2EC24
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTOL 
      Height          =   480
      Left            =   120
      Picture         =   "frmGame.frx":2F866
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIL 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":304A8
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTIL 
      Height          =   480
      Left            =   1560
      Picture         =   "frmGame.frx":310EA
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBIR 
      Height          =   480
      Left            =   2520
      Picture         =   "frmGame.frx":31D2C
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlack 
      Height          =   480
      Left            =   2040
      Picture         =   "frmGame.frx":3296E
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBush 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":335B0
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   320
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Image imgGrass 
      Height          =   480
      Left            =   3480
      Picture         =   "frmGame.frx":341F2
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSign 
      Height          =   480
      Left            =   1080
      Picture         =   "frmGame.frx":34E34
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'******************************************************
'Andres Zacarias
'Quest Game
'Game Type: RPG
'
'This game is an idea from my favourite game : Zelda64 Ocarina
'of time.
'
'If you have any idea of how to improve the speech and the item finding please
'email me.
'Im really not comenting this thing.
'******************************************************
'******************************************************


Option Explicit


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyUp
    'Key Up.
    PositionMap = Mid(Map(CharY + 2), CharX + 3, 1)
    CharFacing = 1
    Label1 = CharY
    Call CharacterMovements
    If txtSpeech.Visible = True Then txtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharY = CharY - 1
      DrawIt
    End If
    If PositionMap = "#" Then
      Call SelectPlace
      DrawIt
    End If
  Case vbKeyDown
    'Key Down.
    PositionMap = Mid(Map(CharY + 4), CharX + 3, 1)
    CharFacing = 2
    Label1 = CharY
    Call CharacterMovements
    If txtSpeech.Visible = True Then txtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharY = CharY + 1
      DrawIt
    End If
    If PositionMap = "M" Then
      Call SelectPlace
      DrawIt
    End If
  Case vbKeyRight
    'Key Right.
    PositionMap = Mid(Map(CharY + 3), CharX + 4, 1)
    CharFacing = 3
    Label2 = CharX
    Call CharacterMovements
    If txtSpeech.Visible = True Then txtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharX = CharX + 1
      DrawIt
    End If
  Case vbKeyLeft
    'Key Left.
    PositionMap = Mid(Map(CharY + 3), CharX + 2, 1)
    CharFacing = 4
    Label2 = CharX
    Call CharacterMovements
    If txtSpeech.Visible = True Then txtSpeech.Visible = False
    If PositionMap = "G" Or PositionMap = "C" Then
      CharX = CharX - 1
      DrawIt
    End If
  Case vbKeySpace
    'Key Space.
    Call TypeOfItem
    DrawIt
    Call InitializeSpeech
  Case vbKeyZ
    If SpellCut = True Then
      Call Make_Spell_Cut
      DrawIt
    End If
  Case vbKeyEscape
    CloseMidi
    UnHook
  End Select
End Sub

Private Sub Form_Load()
  frmGame.Height = 5175
  frmGame.Width = 4890
  Call InitGame
  Midi = "Town.mid"
  Call InitMusic
End Sub

Public Sub InitGame()
  'Starting the first Map.
  Call Map_A1
  'Load the Speech for the Map.
  Call Speech_A1
  NewLine = Chr(13) + Chr(10)
  'Character Position.
  CharX = 10
  CharY = 7
  'Facing Position.
  CharFacing = 2
  MapLoaded = "A1"
  SpellCut = False
  Wood = 0
  Item = 0
  ItemTownFound1 = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub



Private Sub Timer1_Timer()
  DrawIt
  Timer1.Enabled = False
End Sub

Public Sub DrawIt()
  For Y = -3 To 6
    For X = -3 To 6
      'If the result to Paint is 0 then it will get error.
      'This will prevent this.
      PassToNext = 0
        If Y + CharY + 0 < 1 Then PictureHandler
        If X + CharX + 0 < 1 Then PictureHandler
        If X + CharX + 0 > Len(Map(1)) Then PictureHandler
        If Y + CharY + 0 > 51 Then PictureHandler
      If PassToNext = 0 Then PositionMap = Mid(Map(Y + CharY + 1), (X + CharX + 1), 1)
      'If X = 0 And Y = 0 Then GoTo skip:
      Select Case PositionMap
      Case Is = "G" 'Grass
        PaintPicture imgGrass.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "B" 'Bush
        PaintPicture imgBush.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "<" 'Weed
        PaintPicture imgWeed.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "0" 'Cero Sign
        PaintPicture imgSign.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "Q" 'Water
        PaintPicture imgTOL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "q" 'grass
        PaintPicture imgTOL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "A" 'Water
        PaintPicture imgBOL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "a" 'grass
        PaintPicture imgBOL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "W" 'Water
        PaintPicture imgTOR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "w" 'grass
        PaintPicture imgTOR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "S" 'Water
        PaintPicture imgBOR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "s" 'grass
        PaintPicture imgBOR2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "E" 'Border Left water
        PaintPicture ImgIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "e" 'Border Left grass
        PaintPicture ImgIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "R" 'Border Right water
        PaintPicture ImgIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "r" 'Border Right grass
        PaintPicture ImgIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "D" 'Border Top water
        PaintPicture ImgIT.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "d" 'Border Top grass
        PaintPicture ImgIT2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "F" 'Border Bottom water
        PaintPicture ImgIB.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "f" 'Border Bottom grass
        PaintPicture ImgIB2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "T" 'Border Bottom water
        PaintPicture ImgTIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "t" 'Border Bottom grass
        PaintPicture ImgTIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "Y" 'Border Bottom water
        PaintPicture ImgTIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "y" 'Border Bottom grass
        PaintPicture ImgTIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "V" 'Border Bottom water
        PaintPicture ImgBIL.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "v" 'Border Bottom grass
        PaintPicture ImgBIL2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "H" 'Border Bottom water
        PaintPicture ImgBIR.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "h" 'Border Bottom grass
        PaintPicture ImgBIR2.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "U" 'Water
        PaintPicture ImgTIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "u" 'grass
        PaintPicture ImgIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "I" 'Water
        PaintPicture ImgTIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "i" 'grass
        PaintPicture ImgIB3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "J" 'Water
        PaintPicture ImgBIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "j" 'grass
        PaintPicture ImgIL3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "K" 'Water
        PaintPicture ImgBIR3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "k" 'grass
        PaintPicture ImgIT3.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "," 'grass
        PaintPicture imgHouseExitB1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ";" 'grass
        PaintPicture imgHouseExitB2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "M" 'grass
        PaintPicture imgHouseExit.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "-" 'Water
        PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "_" 'Nothing
        PaintPicture imgNothing.Picture, (X + 3) * 32, (Y + 3) * 32
      
      Case Is = "Z" 'Wall bottom
        PaintPicture imgWallBottom.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "X" 'Wall top
        PaintPicture imgWallTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "C" 'Floor Blue 1
        PaintPicture imgFloor1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "z" 'Floor Tronco
        PaintPicture imgWall2.Picture, (X + 3) * 32, (Y + 3) * 32
            
      Case Is = "1" 'Laddy
        PaintPicture imgCharWomen1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "2" 'Boy 1
        PaintPicture imgCharBoy1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "3" 'Boy 2
        PaintPicture imgCharBoy2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "4" 'Good Wizard
        PaintPicture imgCharGoodWizard.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "5" 'Bad Wizard
        PaintPicture imgCharBadWizard.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "6" 'Laddy
        PaintPicture imgCharWomen2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "7" 'Laddy
        PaintPicture imgCharWomen3.Picture, (X + 3) * 32, (Y + 3) * 32
      
      
      Case Is = "/"
        PaintPicture imgBlueTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "\"
        PaintPicture imgRedTop.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "*"
        PaintPicture imgWindow1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "#" '3
        PaintPicture imgDoor1.Picture, (X + 3) * 32, (Y + 3) * 32
    
      Case Is = "!" 'Jar
        PaintPicture imgJar.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "(" 'Table
        PaintPicture imgTable1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = ")" 'Table2
        PaintPicture imgTable2.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "[" 'Bed
        PaintPicture imgBed1.Picture, (X + 3) * 32, (Y + 3) * 32
      Case Is = "]" 'Bed2
        PaintPicture imgBed2.Picture, (X + 3) * 32, (Y + 3) * 32
    
    
      End Select
skip:
    Next
  Next
  

  'Character Movements.
  Select Case CharFacing
  Case Is = 1
    PaintPicture imgUpChar.Picture, 5 * 32, 5 * 32
  Case Is = 2
    PaintPicture imgDownChar.Picture, 5 * 32, 5 * 32
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * 32, 5 * 32
  Case Is = 4
    PaintPicture imgLeftChar.Picture, 5 * 32, 5 * 32
  Case Is = 5
    PaintPicture imgCarryChar.Picture, 5 * 32, 5 * 32
  End Select
  
  Select Case Item
  Case Is = 1
    PaintPicture imgItem1.Picture, 5 * 32, 4 * 32
    Item = 0
  Case Is = 2
    PaintPicture imgItem2.Picture, 5 * 32, 4 * 32
    Item = 0
  End Select

End Sub

Public Sub PictureHandler()
  PassToNext = 1
  PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
End Sub

Public Sub CharacterMovements()
  'Character Movements.
  Select Case CharFacing
  Case Is = 1
    PaintPicture imgUpChar.Picture, 5 * 32, 5 * 32
  Case Is = 2
    PaintPicture imgDownChar.Picture, 5 * 32, 5 * 32
  Case Is = 3
    PaintPicture imgRightChar.Picture, 5 * 32, 5 * 32
  Case Is = 4
    PaintPicture imgLeftChar.Picture, 5 * 32, 5 * 32
  End Select
End Sub

