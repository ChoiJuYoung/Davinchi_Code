VERSION 5.00
Begin VB.Form FrmPick 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "�����Ͻ� ����� ������ �����ּ���."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label lblB 
      BackStyle       =   0  '����
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblW 
      BackStyle       =   0  '����
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Image ImgB 
      Height          =   2295
      Left            =   2520
      Picture         =   "FrmPick.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image ImgW 
      Height          =   2295
      Left            =   120
      Picture         =   "FrmPick.frx":2154
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FrmPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
FrmGame.TimCardSort.Enabled = False
lblW = "���� �Ͼ�� ��� �� : " & RestW
lblB = "���� ������ ��� �� : " & RestB
End Sub

Private Sub ImgB_Click()
If RestB > 0 Then
    Do Until (CardBool(ForPick) = True) And (ForPick Mod 2 = 0)
        ForPick = ((23 * Rnd) + 0)
    Loop
    Placard(Turn, PlaCardVal(Turn) + 1) = ForPick
    PlaCardVal(Turn) = PlaCardVal(Turn) + 1
    CardBool(ForPick) = False
    LastPick = ForPick
    RestB = RestB - 1
    FrmGame.Enabled = True
    FrmGame.TimCardSort.Enabled = True
    FrmGame.TimCardOp.Enabled = True
    FrmGame.CmdCardSort.Enabled = False
    Unload Me
Else
    MsgBox "������ ī�尡 �������� �ʽ��ϴ�."
End If
End Sub

Private Sub ImgW_Click()
If RestW > 0 Then
    Do Until (CardBool(ForPick) = True) And (ForPick Mod 2 = 1)
        ForPick = ((23 * Rnd) + 0)
    Loop
    Placard(Turn, PlaCardVal(Turn) + 1) = ForPick
    PlaCardVal(Turn) = PlaCardVal(Turn) + 1
    CardBool(ForPick) = False
    LastPick = ForPick
    RestW = RestW - 1
    FrmGame.Enabled = True
    FrmGame.TimCardSort.Enabled = True
    FrmGame.TimCardOp.Enabled = True
    FrmGame.CmdCardSort.Enabled = False
    Unload Me
Else
    MsgBox "�Ͼ�� ī�尡 �������� �ʽ��ϴ�."
End If
End Sub
