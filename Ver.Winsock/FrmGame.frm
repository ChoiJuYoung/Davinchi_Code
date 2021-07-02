VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "DavinChi Code By [Tos2]ChocoTea, Marine & MDV // 0.8Version"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleWidth      =   15240
   StartUpPosition =   2  '화면 가운데
   Begin Davinchi_Code.jcbutton CmdMusic 
      Height          =   735
      Left            =   960
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
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
      Caption         =   "Music"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox TxtChat 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   10560
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox TxtAllChat 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   23
      Top             =   6840
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox TxtLog 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox LstPlayer 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   2760
      TabIndex        =   20
      Top             =   2280
      Width           =   1095
   End
   Begin Davinchi_Code.jcbutton CmdPass 
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   3000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
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
      BackColor       =   12648384
      Caption         =   "Pass"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Timer SockCheck 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7920
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   3120
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimCardOp 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6960
      Top             =   1680
   End
   Begin Davinchi_Code.jcbutton CmdCardSort 
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648384
      Caption         =   "Start"
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.Timer TimCardSort 
      Interval        =   10
      Left            =   7440
      Top             =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "게임 대기자s"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label NickName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   19
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label NickName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   18
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label NickName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "안녕하세요"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   17
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label NickName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "안녕하세요"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   16
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblO 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   14
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   13
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   12
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   11
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   10
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   9
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   8
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   7
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   6
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   5
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   4
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   3
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   2
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla2 
      Height          =   750
      Index           =   1
      Left            =   14400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   14
      Left            =   120
      Top             =   11040
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   13
      Left            =   120
      Top             =   10200
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   12
      Left            =   120
      Top             =   9360
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   11
      Left            =   120
      Top             =   8520
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   10
      Left            =   120
      Top             =   7680
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   9
      Left            =   120
      Top             =   6840
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   8
      Left            =   120
      Top             =   6000
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   7
      Left            =   120
      Top             =   5160
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   6
      Left            =   120
      Top             =   4320
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   5
      Left            =   120
      Top             =   3480
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   4
      Left            =   120
      Top             =   2640
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   3
      Left            =   120
      Top             =   1800
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   2
      Left            =   120
      Top             =   960
      Width           =   750
   End
   Begin VB.Image ImgPla4 
      Height          =   750
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   14
      Left            =   12600
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   13
      Left            =   11760
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   12
      Left            =   10920
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   11
      Left            =   10080
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   10
      Left            =   9240
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   9
      Left            =   8400
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   8
      Left            =   7560
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   7
      Left            =   6720
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   6
      Left            =   5880
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   5
      Left            =   5040
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   4
      Left            =   4200
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   3
      Left            =   3360
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   2
      Left            =   2520
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla3 
      Height          =   750
      Index           =   1
      Left            =   1680
      Top             =   120
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   14
      Left            =   12720
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   13
      Left            =   11880
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   12
      Left            =   11040
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   11
      Left            =   10200
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   10
      Left            =   9360
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   9
      Left            =   8520
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   8
      Left            =   7680
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   7
      Left            =   6840
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   6
      Left            =   6000
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   5
      Left            =   5160
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   4
      Left            =   4320
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   3
      Left            =   3480
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   2
      Left            =   2640
      Top             =   6720
      Width           =   750
   End
   Begin VB.Image ImgPla1 
      Height          =   750
      Index           =   1
      Left            =   1800
      Top             =   6720
      Width           =   750
   End
End
Attribute VB_Name = "FrmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Function ImgArr(a As Integer)
For k = 1 To 14
    If a = 2 Then
        ImgPla2(k).Left = ImgPla3(k).Left
        ImgPla2(k).Top = ImgPla3(k).Top
        NickName(2).Left = NickName(3).Left
        NickName(2).Top = NickName(3).Top
    ElseIf a = 3 Then
        ImgPla3(k).Left = ImgPla4(k).Left
        ImgPla3(k).Top = ImgPla4(k).Top
        NickName(3).Left = NickName(4).Left
        NickName(3).Top = NickName(4).Top
        ImgPla3(k).Visible = True
    Else
        ImgPla3(k).Visible = True
        ImgPla4(k).Visible = True
    End If
Next
End Function

Private Sub CmdCardSort_Click()

If Star = False Then
    FrmMode.Show
'    Turn = 0 '턴 넘길 때 turn++후에 turn / 4의 나머지로 넘김
'    For i = 0 To 3
'        For j = 1 To 14
'            PlacardB(i, j) = False
'        Next
'    Next
'    Randomize ForPick
'    For i = 0 To 25
'        CardBool(i) = True
'    Next
'
'    For j = 0 To 3
'        For i = 1 To 4
'            ForPick = ((23 * Rnd) + 0)
'            Do Until CardBool(ForPick) = True
'                ForPick = ((23 * Rnd) + 0)
'            Loop
'        Placard(j, i) = ForPick
'        CardBool(ForPick) = False
'        If ForPick Mod 2 = 0 Then
'            RestB = RestB - 1
'        Else
'            RestW = RestW - 1
'        End If
'        Next
'        PlaCardVal(j) = 4
'    Next
ElseIf TurnPage = "CARDGET" Then
    If Turn = MyNum Then
        Me.Enabled = False
        FrmPick.Show
    Else
        MsgBox "자신의 차례가 아닙니다."
    End If
Else
    If Turn = MyNum Then
        MsgBox "카드를 더이상 뽑으실 수 없습니다."
    Else
        MsgBox "자신의 차례가 아닙니다."
    End If
End If
End Sub

Private Sub CmdMusic_Click()
FrmMusic.Show
Me.Enabled = False
End Sub

Private Sub CmdPass_Click()
If iscanpass = "true" Then
    Sock.SendData "DC%pass%" & MyNum
Else
    MsgBox "패스할 수 없습니다."
End If
End Sub

Private Sub Command1_Click()
MsgBox PlayerNick(0) & " / " & PlayerNick(1)
End Sub

Private Sub Form_Load()
Sock.Connect "muxacarin.iptime.org", 7788

For i = 1 To 14
    ImgPla3(i).Visible = False
    ImgPla4(i).Visible = False
Next
NickName(1).Visible = False
NickName(2).Visible = False
NickName(3).Visible = False
NickName(4).Visible = False


Dim WGap As Integer, HGap As Integer
WGap = 90 * (XSize / 1280)
HGap = 90 * (YSize / 1024)

For i = 1 To 14
    ImgPla1(i).Top = ImgPla4(14).Top
    ImgPla2(i).Top = ImgPla4(15 - i).Top
    ImgPla3(i).Left = ImgPla1(15 - i).Left
Next

'크기 재배치
TxtAllChat.Width = Me.Width / 2
TxtChat.Width = Me.Width / 2
TxtAllChat.Height = TxtAllChat.Height * YSize / 1024
TxtChat.Height = TxtChat.Height * YSize / 1024
Me.Height = Me.Height * YSize / 1024 + 2 * HGap
Me.Width = Me.Width * XSize / 1280 + 2 * WGap
CmdCardSort.Height = CmdCardSort.Height * YSize / 1024
CmdCardSort.Width = CmdCardSort.Width * XSize / 1280
CmdPass.Height = CmdCardSort.Height
CmdPass.Width = CmdCardSort.Width
TxtLog.Width = TxtLog.Width * XSize / 1280
TxtLog.Height = TxtLog.Height * YSize / 1024
LstPlayer.Height = LstPlayer.Height * YSize / 1024
LstPlayer.Width = LstPlayer.Width * XSize / 1280
For TurnHelp = 1 To 14
    ImgPla1(TurnHelp).Stretch = True
    ImgPla2(TurnHelp).Stretch = True
    ImgPla3(TurnHelp).Stretch = True
    ImgPla4(TurnHelp).Stretch = True
    ImgPla1(TurnHelp).Height = ImgPla1(TurnHelp).Height * YSize / 1024
    ImgPla1(TurnHelp).Width = ImgPla1(TurnHelp).Width * XSize / 1280
    ImgPla2(TurnHelp).Height = ImgPla2(TurnHelp).Height * YSize / 1024
    ImgPla2(TurnHelp).Width = ImgPla2(TurnHelp).Width * XSize / 1280
    ImgPla3(TurnHelp).Height = ImgPla3(TurnHelp).Height * YSize / 1024
    ImgPla3(TurnHelp).Width = ImgPla3(TurnHelp).Width * XSize / 1280
    ImgPla4(TurnHelp).Height = ImgPla4(TurnHelp).Height * YSize / 1024
    ImgPla4(TurnHelp).Width = ImgPla4(TurnHelp).Width * XSize / 1280
Next

ImgPla3(14).Left = ImgPla4(1).Left + 2 * ImgPla4(1).Width + WGap

For TurnHelp = 1 To 13
    ImgPla4(TurnHelp + 1).Top = ImgPla4(TurnHelp).Top + ImgPla4(TurnHelp).Height + HGap
    ImgPla3(14 - TurnHelp).Left = ImgPla3(14 - (TurnHelp - 1)).Left + ImgPla3(14 - (TurnHelp - 1)).Width + WGap
Next
For TurnHelp = 1 To 14
    ImgPla2(15 - TurnHelp).Top = ImgPla4(TurnHelp).Top
    ImgPla1(15 - TurnHelp).Left = ImgPla3(TurnHelp).Left
    ImgPla2(TurnHelp).Left = ImgPla3(1).Left + 2 * ImgPla3(1).Width + WGap
    ImgPla1(TurnHelp).Top = ImgPla4(14).Top
Next


For i = 1 To 14
    lblO(i).Width = lblO(i).Width * (XSize / 1280)
    lblO(i).Height = lblO(i).Height * (YSize / 1024)
    lblO(i).Left = ImgPla1(i).Left
    lblO(i).Top = ImgPla1(i).Top + ImgPla1(i).Height + 50
    lblO(i).Visible = False
    If lblO(i).Height < 170 Then
        lblO(i).FontSize = 9
    End If
Next

NickName(1).Left = (Me.Width - NickName(1).Width) / 2
NickName(1).Top = ImgPla1(1).Top - (ImgPla2(1).Top - (ImgPla2(2).Top + ImgPla2(2).Height)) * 3
NickName(3).Left = (Me.Width - NickName(1).Width) / 2
NickName(3).Top = (ImgPla3(1).Top + ImgPla3(1).Height) + (ImgPla2(1).Top - (ImgPla2(2).Top + ImgPla2(2).Height))
NickName(2).Top = Me.Height / 2 - NickName(2).Height / 2
NickName(4).Top = NickName(2).Top
NickName(2).Left = ImgPla2(1).Left - NickName(2).Width - 75 * (XSize / 1280)
NickName(4).Left = ImgPla4(1).Left + NickName(4).Width + 75 * (XSize / 1280)

TxtLog.Left = (Me.Width - TxtLog.Width) / 2
TxtLog.Top = NickName(3).Top + NickName(3).Height + ImgPla2(1).Top - (ImgPla2(2).Top + ImgPla2(2).Height)

CmdCardSort.Left = Me.Width / 2 - CmdCardSort.Width / 2
CmdPass.Left = CmdCardSort.Left
CmdCardSort.Top = Me.Height / 2 - CmdCardSort.Height
CmdPass.Top = CmdCardSort.Top + CmdCardSort.Height + (ImgPla2(1).Top - (ImgPla2(2).Top + ImgPla2(2).Height))

TxtAllChat.Top = CmdPass.Top + CmdPass.Height + (ImgPla2(1).Top - (ImgPla2(2).Top + ImgPla2(2).Height))
TxtChat.Top = TxtAllChat.Top + TxtAllChat.Height + (ImgPla2(1).Top - (ImgPla2(2).Top + ImgPla2(2).Height))
TxtAllChat.Left = (Me.Width - TxtAllChat.Width) / 2
TxtChat.Left = TxtAllChat.Left

Star = False
ImgPla4(14).Top = ImgPla4(13).Top + ImgPla4(13).Top - ImgPla4(12).Top

LstPlayer.Top = (CmdPass.Top + CmdCardSort.Top + CmdPass.Height) / 2 - LstPlayer.Height / 2
LstPlayer.Left = CmdPass.Left - LstPlayer.Width - (ImgPla1(2).Left - (ImgPla1(1).Left + ImgPla1(1).Width))
Label1.Top = LstPlayer.Top - 360 * (YSize / 1024)
Label1.Left = LstPlayer.Left

End Sub


Private Sub Form_Unload(Cancel As Integer)
Sock.Close
End Sub

Private Sub ImgPla2_Click(Index As Integer)
If Turn = MyNum Then
    If TurnPage = "CARDGET" Then
        MsgBox "카드를 뽑아주세요."
    Else
        If Placard((MyNum + 1) Mod PlayerNum, Index) <> "" Then
            TheAnswerCo = Placard((MyNum + 1) Mod PlayerNum, Index)
            AnswerClickIndex = Index - 1
            FrmAns.Show
            PickPla = 1
            Me.Enabled = False
        Else
            MsgBox "그 곳에는 상대방의 카드가 없습니다."
        End If
    End If
Else
    MsgBox "자신의 차례가 아닙니다."
End If
End Sub

Private Sub ImgPla3_Click(Index As Integer)
If Turn = MyNum Then
    If TurnPage = "CARDGET" Then
        MsgBox "카드를 뽑아주세요."
    Else
        If Placard((MyNum + 2) Mod PlayerNum, Index) <> "" Then
            TheAnswerCo = Placard((MyNum + 2) Mod PlayerNum, Index)
            AnswerClickIndex = Index - 1
            FrmAns.Show
            PickPla = 2
            Me.Enabled = False
        Else
            MsgBox "그 곳에는 상대방의 카드가 없습니다."
        End If
    End If
Else
    MsgBox "자신의 차례가 아닙니다."
End If
End Sub

Private Sub ImgPla4_Click(Index As Integer)
If Turn = MyNum Then
    If TurnPage = "CARDGET" Then
        MsgBox "카드를 뽑아주세요."
    Else
        If Placard((MyNum + 3) Mod PlayerNum, Index) <> "" Then
            TheAnswerCo = Placard((MyNum + 3) Mod PlayerNum, Index)
            AnswerClickIndex = Index - 1
            FrmAns.Show
            PickPla = 3
            Me.Enabled = False
        Else
            MsgBox "그 곳에는 상대방의 카드가 없습니다."
        End If
    End If
Else
    MsgBox "자신의 차례가 아닙니다."
End If
End Sub


Private Sub Sock_Connect()
Sock.SendData "DC%nick%" & Nick
SockCheck.Enabled = True
'BGM.Movie = App.Path & "\BackGroundMusic.swf"
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
Dim chat As Boolean
chat = False


Dim RE As String 'Data저장
Sock.GetData RE

Dim RES() As String 'Data 분리용 변수
Dim Ubnd As Integer 'DC 갯수 계산
Dim RESS() As String 'Data 재분리용 변수
Dim UbndS As Integer '%갯수 계산

RES() = Split(RE, "!^#") 'Data를 DC를 기준으로 분리
Ubnd = UBound(RES)
MsgBox (RE)


For i = 0 To Ubnd - 1
    If Len(RES(i)) >= 1 Then
'        Text1 = Text1 & RES(i)
        RESS() = Split(RES(i), "%")
        UbndS = UBound(RESS)
        If RESS(1) = "start" Then
            CmdMusic.Visible = True
            'bgm.movie = App.Path & "\bgm\BackGroundMusic 01.swf"
            TxtLog.Visible = True
            TxtAllChat.Visible = True
            TxtChat.Visible = True
            PlayerNum = RESS(2)
            CmdCardSort.Enabled = True
            CmdCardSort.Caption = "카드 뽑기"
            Star = True
            Call ImgArr(PlayerNum)
            Label1.Visible = False
            LstPlayer.Visible = False
            NickName(1).Visible = True
            NickName(2).Visible = True
            If PlayerNum = 3 Then
                NickName(3).Visible = True
            ElseIf PlayerNum = 4 Then
                NickName(3).Visible = True
                NickName(4).Visible = True
            End If

        ElseIf RESS(1) = "myindex" Then
            MyNum = Val(RESS(2))
            NickName(1) = PlayerNick(MyNum)
            NickName(2) = PlayerNick((MyNum + 1) Mod PlayerNum)
            If PlayerNum = 3 Then
                NickName(3) = PlayerNick((MyNum + 2) Mod PlayerNum)
            ElseIf PlayerNum = 4 Then
                NickName(3) = PlayerNick((MyNum + 2) Mod PlayerNum)
                NickName(4) = PlayerNick((MyNum + 3) Mod PlayerNum)
            End If
        ElseIf RESS(1) = "myturn" Or RESS(1) = "otherturn" Then
            Turn = Val(RESS(2)) Mod PlayerNum
        ElseIf RESS(1) = "turnindex" Then
            TurnPage = RESS(2)
        ElseIf RESS(1) = "remainBlack" Then
            RestB = RESS(2)
        ElseIf RESS(1) = "remainWhite" Then
            RestW = RESS(2)
        ElseIf RESS(1) = "card" Then
            PlaCardVal(RESS(2)) = UbndS - 2
            For j = 3 To UbndS
                Placard(RESS(2), j - 2) = RESS(j)
            Next
        ElseIf RESS(1) = "dead" Then
            PlayerLive(RESS(2)) = False
        ElseIf RESS(1) = "iscanpass" Then
            iscanpass = RESS(2)
        ElseIf RESS(1) = "chat" Then
            TxtAllChat = TxtAllChat & PlayerNick(RESS(2)) & " : " & RESS(3) & vbCrLf
            chat = True
        ElseIf RESS(1) = "end" Then
            SockCheck.Enabled = False
            Sock.Close
            If MyNum = RESS(2) Then
                MsgBox "승리하셨습니다."
            Else
                MsgBox "패배하셨습니다."
            End If
            End
        ElseIf RESS(1) = "error" Then
            MsgBox "인원이 맞지 않습니다."
            GoTo error:
        ElseIf RESS(1) = "cardgot" Then
            TxtLog = TxtLog & vbCrLf & PlayerNick(RESS(2)) & "님이 " & RESS(3) & "카드를 가져오셨습니다. 위치는 " & RESS(4) + 1 & "번째입니다." & vbCrLf
        ElseIf RESS(1) = "cardchoice" Then
            TxtLog = TxtLog & vbCrLf & PlayerNick(RESS(2)) & "님이 " & PlayerNick((Val(RESS(2)) + Val(RESS(3))) Mod PlayerNum) & "님의 " & RESS(4) + 1 & "번째 카드가 " & RESS(5) & "색깔의 " & RESS(6) & "이라고 공격하셨습니다." & vbCrLf
        ElseIf RESS(1) = "nick" Then
            LstPlayer.Clear
            For j = 2 To UbndS
                LstPlayer.AddItem (RESS(j))
            Next
        ElseIf RESS(1) = "inxnick" Then
            PlayerNick(RESS(2)) = RESS(3)
        End If
    End If
Next

If Star = True Then

    If PlayerNum = 4 Then
        If Turn Mod PlayerNum = (MyNum) Mod PlayerNum Then
            TurnDir = "▼"
        ElseIf Turn Mod PlayerNum = (MyNum + 1) Mod PlayerNum Then
            TurnDir = "▶"
        ElseIf Turn Mod PlayerNum = (MyNum + 2) Mod PlayerNum Then
            TurnDir = "▲"
        ElseIf Turn Mod PlayerNum = (MyNum + 3) Mod PlayerNum Then
            TurnDir = "◀"
        End If
    ElseIf PlayerNum = 3 Then
        If Turn Mod PlayerNum = (MyNum) Mod PlayerNum Then
            TurnDir = "▼"
        ElseIf Turn Mod PlayerNum = (MyNum + 1) Mod PlayerNum Then
            TurnDir = "▶"
        ElseIf Turn Mod PlayerNum = (MyNum + 2) Mod PlayerNum Then
            TurnDir = "◀"
        End If
    ElseIf PlayerNum = 2 Then
        If Turn Mod PlayerNum = (MyNum) Mod PlayerNum Then
            TurnDir = "▼"
        ElseIf Turn Mod PlayerNum = (MyNum + 1) Mod PlayerNum Then
            TurnDir = "▲"
        End If
    End If
    
    If TurnPage = "CARDGET" Then
        CmdCardSort.Caption = "카드를 뽑아주세요."
    ElseIf TurnPage = "SELECT" Then
        CmdCardSort.Caption = "상대의 카드를 맞춰주세요."
    End If
    
    If (RestW + RestB) <= 0 Then
        CmdCardSort.Caption = "더이상 카드가 남지 않았습니다."
        CmdCardSort.Enabled = False
    End If
    
    CmdCardSort.Caption = CmdCardSort.Caption & vbCrLf & "Player = " & TurnDir
    If chat = False Then
        TimCardOp.Enabled = True
    End If
End If

Exit Sub

error:
End Sub

Private Sub SockCheck_Timer()
If Sock.State <> 7 Then
    MsgBox "서버와의 연결이 끊겼습니다. 게임을 종료합니다."
    End
End If
End Sub


Private Sub TimCardOp_Timer()
TimCardOp.Enabled = False
For k = 1 To 14
    ImgPla1(k) = Nothing
    ImgPla2(k) = Nothing
    ImgPla3(k) = Nothing
    ImgPla4(k) = Nothing
Next

For j = 1 To 14
    lblO(j) = "None"
    lblO(j).ForeColor = RGB(0, 0, 0)
    lblO(j).Visible = False
Next
DoEvents

For k = 0 To PlayerNum - 1
    For j = 1 To PlaCardVal((MyNum + k) Mod PlayerNum)
        If Len(Placard((MyNum + k) Mod PlayerNum, j)) > 1 Then
            If InStr(Placard((MyNum + k) Mod PlayerNum, j), "O") <> 0 Then
                Placard((MyNum + k) Mod PlayerNum, j) = Left(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1))
                lblO(j) = "Done"
                lblO(j).ForeColor = RGB(96, 255, 128)
            End If
            If Left(Placard((MyNum + k) Mod PlayerNum, j), 1) = "B" Then
                If k = 0 Then
                    ImgPla1(j) = LoadPicture(App.Path & "\img\" & Val(Right(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1)) * 2) & ".jpg")
                    lblO(j).Visible = True
                ElseIf k = 1 Then
                    ImgPla2(j) = LoadPicture(App.Path & "\img\" & Val(Right(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1)) * 2) & ".jpg")
                ElseIf k = 2 Then
                    ImgPla3(j) = LoadPicture(App.Path & "\img\" & Val(Right(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1)) * 2) & ".jpg")
                ElseIf k = 3 Then
                    ImgPla4(j) = LoadPicture(App.Path & "\img\" & Val(Right(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1)) * 2) & ".jpg")
                End If
            ElseIf Left(Placard((MyNum + k) Mod PlayerNum, j), 1) = "W" Then
                If k = 0 Then
                    lblO(j).Visible = True
                    ImgPla1(j) = LoadPicture(App.Path & "\img\" & Val(Right(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1)) * 2 + 1) & ".jpg")
                ElseIf k = 1 Then
                    ImgPla2(j) = LoadPicture(App.Path & "\img\" & Val(Right(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1)) * 2 + 1) & ".jpg")
                ElseIf k = 2 Then
                    ImgPla3(j) = LoadPicture(App.Path & "\img\" & Val(Right(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1)) * 2 + 1) & ".jpg")
                ElseIf k = 3 Then
                    ImgPla4(j) = LoadPicture(App.Path & "\img\" & Val(Right(Placard((MyNum + k) Mod PlayerNum, j), Val(Len(Placard((MyNum + k) Mod PlayerNum, j)) - 1)) * 2 + 1) & ".jpg")
                End If
            End If
        Else
            If (Placard((MyNum + k) Mod PlayerNum, j) = "B") Then
                If k = 0 Then
                    ImgPla1(j) = LoadPicture(App.Path & "\img\B.jpg")
                ElseIf k = 1 Then
                    ImgPla2(j) = LoadPicture(App.Path & "\img\B.jpg")
                ElseIf k = 2 Then
                    ImgPla3(j) = LoadPicture(App.Path & "\img\B.jpg")
                ElseIf k = 3 Then
                    ImgPla4(j) = LoadPicture(App.Path & "\img\B.jpg")
                End If
            ElseIf (Placard((MyNum + k) Mod PlayerNum, j) = "W") Then
                If k = 0 Then
                    ImgPla1(j) = LoadPicture(App.Path & "\img\W.jpg")
                ElseIf k = 1 Then
                    ImgPla2(j) = LoadPicture(App.Path & "\img\W.jpg")
                ElseIf k = 2 Then
                    ImgPla3(j) = LoadPicture(App.Path & "\img\W.jpg")
                ElseIf k = 3 Then
                    ImgPla4(j) = LoadPicture(App.Path & "\img\W.jpg")
                End If
            End If
        End If
    Next
Next


'For k = 0 To 3
'    For j = 1 To 4
'        For i = 1 To PlaCardVal((Turn + k) Mod 4) - 1
'            If Placard((Turn + k) Mod 4, i) > Placard((Turn + k) Mod 4, i + 1) Then
'                num = Placard((Turn + k) Mod 4, i)
'                Placard((Turn + k) Mod 4, i) = Placard((Turn + k) Mod 4, i + 1)
'                Placard((Turn + k) Mod 4, i + 1) = num
'            End If
'        Next
'    Next
'Next
'
'For i = 1 To PlaCardVal(Turn)
'    ImgPla1(i) = LoadPicture(App.Path & "\img\" & Placard(Turn, i) & ".jpg")
'Next
'
'For j = 1 To 3
'    For i = 1 To PlaCardVal((Turn + j) Mod 4)
'        If j = 1 Then
'            If PlacardB((Turn + j) Mod 4, i) = True Then
'                ImgPla2(i) = LoadPicture(App.Path & "\img\" & Placard((Turn + j) Mod 4, i) & ".jpg")
'            Else
'                If (Placard((Turn + j) Mod 4, i) Mod 2) = 0 Then
'                    ImgPla2(i) = LoadPicture(App.Path & "\img\B.jpg")
'                Else
'                    ImgPla2(i) = LoadPicture(App.Path & "\img\W.jpg")
'                End If
'            End If
'        ElseIf j = 2 Then
'            If PlacardB((Turn + j) Mod 4, i) = True Then
'                ImgPla3(i) = LoadPicture(App.Path & "\img\" & Placard((Turn + j) Mod 4, i) & ".jpg")
'            Else
'                If (Placard((Turn + j) Mod 4, i) Mod 2) = 0 Then
'                    ImgPla3(i) = LoadPicture(App.Path & "\img\B.jpg")
'                Else
'                    ImgPla3(i) = LoadPicture(App.Path & "\img\W.jpg")
'                End If
'            End If
'        Else
'            If PlacardB((Turn + j) Mod 4, i) = True Then
'                ImgPla4(i) = LoadPicture(App.Path & "\img\" & Placard((Turn + j) Mod 4, i) & ".jpg")
'            Else
'                If (Placard((Turn + j) Mod 4, i) Mod 2) = 0 Then
'                    ImgPla4(i) = LoadPicture(App.Path & "\img\B.jpg")
'                Else
'                    ImgPla4(i) = LoadPicture(App.Path & "\img\W.jpg")
'                End If
'            End If
'        End If
'    Next
'Next

End Sub


Private Sub TxtAllChat_Change()
TxtAllChat.SelLength = Len(TxtAllChat.Text)
End Sub

Private Sub TxtChat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Sock.SendData "DC%chat%" & MyNum & "%" & TxtChat
    TxtChat = ""
End If
End Sub

Private Sub TxtLog_Change()
TxtLog.SelLength = Len(TxtLog.Text)
End Sub
