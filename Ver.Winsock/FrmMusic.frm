VERSION 5.00
Begin VB.Form FrmMusic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin Davinchi_Code.jcbutton CmdChange 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Music Change"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin Davinchi_Code.jcbutton CmdOnOff 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Music On / Off"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
End
Attribute VB_Name = "FrmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdChange_Click()
If MusicOpt = 1 Then
    MusicOpt = 2
Else
    MusicOpt = 1
End If

FrmGame.bgm.playing = False
FrmGame.bgm.movie = App.Path & "\bgm\BackGroundMusic 0" & MusicOpt & ".swf"
FrmGame.bgm.playing = True
FrmGame.Enabled = True
Unload Me
End Sub

Private Sub CmdOnOff_Click()
FrmGame.bgm.playing = Not FrmGame.bgm.playing
FrmGame.Enabled = True
Unload Me
End Sub
