VERSION 5.00
Begin VB.Form FrmAns 
   BackColor       =   &H00000000&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Davinchi_Code.jcbutton CmdAns 
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "입력"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox TxtAns 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmAns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAns_Click()
If TheAnswer = TxtAns Then
    MsgBox "정답입니다."
    For TurnHelp = 1 To 4
        If TurnHelp = 4 Then
            MsgBox "승리하셨습니다."
            End
        End If
        If PlayerLive((Turn + TurnHelp) Mod 4) = True Then
            Exit For
        End If
    Next
    PlaRestCard((Turn + PickPla) Mod 4) = PlaRestCard((Turn + PickPla) Mod 4) - 1
    If PlaRestCard((Turn + PickPla) Mod 4) <= 0 Then
        PlayerLive((Turn + PickPla) Mod 4) = False
    End If
    PlacardB((Turn + PickPla) Mod 4, ClickCard) = True
    If MsgBox("계속 맞추시겠습니까?", vbYesNo) = vbYes Then
        PlacardB((Turn + PickPla) Mod 4, ClickCard) = True
        FrmGame.TimCardOp.Enabled = True
    Else
        TurnHelp = 1
        DoEvents
        Do While PlayerLive((Turn + TurnHelp) Mod 4) = False
            TurnHelp = TurnHelp + 1
        Loop
        Turn = (Turn + TurnHelp) Mod 4
        GoPage = 0
            
        If RestW + RestB > 0 Then
            FrmGame.CmdCardSort.Enabled = True
        Else
            FrmGame.CmdCardSort.Caption = "남은 블록이 없습니다."
            GoPage = 1
        End If
    End If
Else
    MsgBox "틀렸습니다."
    For i = 1 To PlaCardVal(Turn)
        If Placard(Turn, i) = LastPick Then
            PlacardB(Turn, i) = True
            Exit For
        End If
    Next
    TurnHelp = 1
    DoEvents
    Do While PlayerLive((Turn + TurnHelp) Mod 4) = False
        TurnHelp = TurnHelp + 1
    Loop
    Turn = (Turn + TurnHelp) Mod 4
    GoPage = 0
    If RestW + RestB > 0 Then
        FrmGame.CmdCardSort.Enabled = True
    Else
        FrmGame.CmdCardSort.Caption = "남은 블록이 없습니다."
        GoPage = 1
    End If
End If

FrmGame.TimCardOp.Enabled = True
FrmGame.Enabled = True
Unload Me
End Sub

Private Sub TxtAns_Click()
TxtAns = ""
End Sub
