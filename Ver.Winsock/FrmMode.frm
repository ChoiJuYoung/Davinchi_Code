VERSION 5.00
Begin VB.Form FrmMode 
   BackColor       =   &H00008000&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin Davinchi_Code.jcbutton CmdNormal 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      BackColor       =   12648384
      Caption         =   "일반 모드"
      ForeColorHover  =   16761087
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   16761087
      TooltipBackColor=   16761087
      ColorScheme     =   3
   End
   Begin Davinchi_Code.jcbutton CmdBadak 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      BackColor       =   12648384
      Caption         =   "바닥 모드"
      ForeColorHover  =   16761087
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      MaskColor       =   16761087
      TooltipBackColor=   16761087
      ColorScheme     =   3
   End
End
Attribute VB_Name = "FrmMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdBadak_Click()
FrmGame.Sock.SendData "DC%start%0"

Unload Me
End Sub

Private Sub CmdNormal_Click()
FrmGame.Sock.SendData "DC%start%1"

Unload Me
End Sub
