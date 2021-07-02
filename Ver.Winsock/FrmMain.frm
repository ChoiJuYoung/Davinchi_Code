VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Made by 오지석, 최주영 in UOS Computer Science."
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9990
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   240
      Top             =   3240
   End
   Begin MSWinsockLib.Winsock CheckSock 
      Left            =   840
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Davinchi_Code.jcbutton CmdExit 
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   5040
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1296
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "정말 아쉽고 마음이 텅 빈 것처럼 느껴지지만 그러한 마음을 뒤로 하고 종료하기"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Davinchi_Code.jcbutton CmdMult 
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   4200
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1296
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16744703
      Caption         =   "여러명이 즐기기 (현재 구현 : 4인용)"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Davinchi_Code.jcbutton CmdSing 
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   3360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1296
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16744576
      Caption         =   "혼자 즐기기"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Server의 상태를 확인중입니다 ..."
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   7815
   End
   Begin VB.Image Label1 
      Height          =   1680
      Left            =   0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   10110
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CheckSock_Connect()
CheckSock.SendData "HELO"
Timer1.Enabled = False
Label2 = "Connection. Checking Server"
End Sub


Private Sub CheckSock_DataArrival(ByVal bytesTotal As Long)
Dim RE As String
CheckSock.GetData RE

If Left(RE, 3) = "250" Then
    CheckSock.SendData "STAT"
ElseIf (Left(RE, 3) = "330") Or (Left(RE, 3) = "400") Then
    CmdMult.Enabled = True
    Label2 = "Server is now on"
    CheckSock.SendData "QUIT"
ElseIf Left(RE, 3) = "340" Then
    CheckSock.SendData "STRT"
ElseIf Left(RE, 3) = "221" Then
    CheckSock.Close
End If
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdMult_Click()
FrmNick.Show
Unload Me

End Sub

Private Sub CmdSing_Click()
MsgBox "아직 지원하지 않는 기능입니다. 죄송합니다 우 ㅅ유"
End Sub

Private Sub Form_Load()
Dim ButtGap As Integer, SizeGap As Integer


XSize = Screen.Width / Screen.TwipsPerPixelX
YSize = Screen.Height / Screen.TwipsPerPixelY
ButtGap = 105 * (YSize / 1024)
SizeGap = 1200 * (YSize / 1024)

'크기 재배치
Me.Height = Me.Height * YSize / 1024 + 2 * ButtGap
Me.Width = Me.Width * XSize / 1280
Label1.Height = Label1.Height * YSize / 1024
Label1.Width = Label1.Width * XSize / 1280
CmdSing.Height = CmdSing.Height * YSize / 1024
CmdSing.Width = CmdSing.Width * XSize / 1280
CmdMult.Height = CmdMult.Height * YSize / 1024
CmdMult.Width = CmdMult.Width * XSize / 1280
CmdExit.Height = CmdExit.Height * YSize / 1024
CmdExit.Width = CmdExit.Width * XSize / 1280
Label2.Width = CmdMult.Width
Label2.Height = Label2.Height * YSize / 1024

Label1.Top = Label1.Top * (YSize / 1024)
CmdSing.Top = Label1.Top + Label1.Height + SizeGap
CmdMult.Top = CmdSing.Top + CmdSing.Height + ButtGap
CmdExit.Top = CmdMult.Top + CmdMult.Height + ButtGap
Label2.Top = (CmdSing.Top - (CmdSing.Top - (Label1.Top + Label1.Height)) / 2) - Label2.Height / 2
Label1.Picture = LoadPicture(App.Path & "\img\Title.jpg")



RestW = 12
RestB = 12
For i = 0 To 3
    For j = 1 To 14
        PlacardB(i, j) = False
    Next
    PlaRestCard(i) = 4
    PlayerLive(i) = True
Next
End Sub

Private Sub Timer1_Timer()
If CheckSock.State = 0 Then
    CheckSock.Connect "muxacarin.iptime.org", 7700
Else
    If Label2 = "Server의 상태를 확인중입니다 ..." Then
        Label2 = "Server의 상태를 확인중입니다 ."
    ElseIf Label2 = "Server의 상태를 확인중입니다 .." Then
        Label2 = "Server의 상태를 확인중입니다 ..."
    Else
        Label2 = "Server의 상태를 확인중입니다 .."
    End If
End If

End Sub
