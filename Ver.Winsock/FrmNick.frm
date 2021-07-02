VERSION 5.00
Begin VB.Form FrmNick 
   BackColor       =   &H00000000&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Davinchi_Code.jcbutton jcbutton1 
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "완료"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "NickName"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "닉네임을 입력해주세요. (1 ~ 5글자)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmNick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub jcbutton1_Click()
If Len(Text1) > 0 And Len(Text1) < 6 Then
    Nick = Text1
    FrmGame.Show
    Unload Me
Else
    MsgBox "닉네임은 한글자 ~ 다섯글자로 해주세요."
End If
End Sub

Private Sub Text1_Change()
If Len(Text1) > 5 Then
    Text1 = Left(Text1, 5)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    jcbutton1_Click
End If
End Sub
