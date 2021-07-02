VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmGame 
   BackColor       =   &H00FFC0FF&
   Caption         =   "DavinChi Code By Hiasen, Marine & MDV"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleWidth      =   15240
   StartUpPosition =   2  '화면 가운데
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
      BackColor       =   16761087
      Caption         =   "패 섞기"
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.Timer TimCardSort 
      Interval        =   10
      Left            =   7440
      Top             =   1680
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

Private Sub CmdCardSort_Click()
Sock.SendData "DC%start"

If Star = False Then
    Turn = 0 '턴 넘길 때 turn++후에 turn / 4의 나머지로 넘김
    For i = 0 To 3
        For j = 1 To 14
            PlacardB(i, j) = False
        Next
    Next
    Randomize ForPick
    For i = 0 To 25
        CardBool(i) = True
    Next
    
    For j = 0 To 3
        For i = 1 To 4
            ForPick = ((23 * Rnd) + 0)
            Do Until CardBool(ForPick) = True
                ForPick = ((23 * Rnd) + 0)
            Loop
        Placard(j, i) = ForPick
        CardBool(ForPick) = False
        If ForPick Mod 2 = 0 Then
            RestB = RestB - 1
        Else
            RestW = RestW - 1
        End If
        Next
        PlaCardVal(j) = 4
    Next
    
    CmdCardSort.Enabled = True
    
    TimCardOp.Enabled = True
    CmdCardSort.Caption = "카드 뽑기"
    Star = True
Else
    Me.Enabled = False
    GoPage = 1
    FrmPick.Show
End If
End Sub

Private Sub Form_Load()
Sock.Connect "127.0.0.1", 7788

Dim WGap As Integer, HGap As Integer
WGap = 90 * (XSize / 1280)
HGap = 90 * (YSize / 1024)

For i = 1 To 14
    ImgPla1(i).Top = ImgPla4(14).Top
    ImgPla2(i).Top = ImgPla4(15 - i).Top
    ImgPla3(i).Left = ImgPla1(15 - i).Left
Next

'크기 재배치
Me.Height = Me.Height * YSize / 1024 + 2 * HGap
Me.Width = Me.Width * XSize / 1280 + 2 * WGap
CmdCardSort.Height = CmdCardSort.Height * YSize / 1024
CmdCardSort.Width = CmdCardSort.Width * XSize / 1280
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

CmdCardSort.Left = Me.Width / 2 - CmdCardSort.Width / 2
CmdCardSort.Top = Me.Height / 2 - CmdCardSort.Height / 2

Star = False
GoPage = 0
ImgPla4(14).Top = ImgPla4(13).Top + ImgPla4(13).Top - ImgPla4(12).Top

End Sub


Private Sub ImgPla2_Click(Index As Integer)
If GoPage = 0 Then
    MsgBox "카드를 뽑아주세요."
Else
    ClickCard = Index
    TheAnswer = Split(Placard((Turn + 1) Mod 4, ClickCard) / 2, ".")(0)
    FrmAns.Show
    PickPla = 1
    Me.Enabled = False
End If
End Sub

Private Sub ImgPla3_Click(Index As Integer)
If GoPage = 0 Then
    MsgBox "카드를 뽑아주세요."
Else
    ClickCard = Index
    TheAnswer = Split(Placard((Turn + 2) Mod 4, ClickCard) / 2, ".")(0)
    FrmAns.Show
    PickPla = 2
    Me.Enabled = False
End If
End Sub

Private Sub ImgPla4_Click(Index As Integer)
If GoPage = 0 Then
    MsgBox "카드를 뽑아주세요."
Else
    ClickCard = Index
    TheAnswer = Split(Placard((Turn + 3) Mod 4, ClickCard) / 2, ".")(0)
    FrmAns.Show
    PickPla = 3
    Me.Enabled = False
End If
End Sub

Private Sub Sock_Connect()
MsgBox "연결 완료"
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
Dim a As String
Sock.GetData a
MsgBox a
End Sub

Private Sub TimCardOp_Timer()
For k = 1 To 14
    ImgPla1(k) = Nothing
    ImgPla2(k) = Nothing
    ImgPla3(k) = Nothing
    ImgPla4(k) = Nothing
Next

For k = 0 To 3
    For j = 1 To 4
        For i = 1 To PlaCardVal((Turn + k) Mod 4) - 1
            If Placard((Turn + k) Mod 4, i) > Placard((Turn + k) Mod 4, i + 1) Then
                num = Placard((Turn + k) Mod 4, i)
                Placard((Turn + k) Mod 4, i) = Placard((Turn + k) Mod 4, i + 1)
                Placard((Turn + k) Mod 4, i + 1) = num
            End If
        Next
    Next
Next

For i = 1 To PlaCardVal(Turn)
    ImgPla1(i) = LoadPicture(App.Path & "\img\" & Placard(Turn, i) & ".jpg")
Next

For j = 1 To 3
    For i = 1 To PlaCardVal((Turn + j) Mod 4)
        If j = 1 Then
            If PlacardB((Turn + j) Mod 4, i) = True Then
                ImgPla2(i) = LoadPicture(App.Path & "\img\" & Placard((Turn + j) Mod 4, i) & ".jpg")
            Else
                If (Placard((Turn + j) Mod 4, i) Mod 2) = 0 Then
                    ImgPla2(i) = LoadPicture(App.Path & "\img\B.jpg")
                Else
                    ImgPla2(i) = LoadPicture(App.Path & "\img\W.jpg")
                End If
            End If
        ElseIf j = 2 Then
            If PlacardB((Turn + j) Mod 4, i) = True Then
                ImgPla3(i) = LoadPicture(App.Path & "\img\" & Placard((Turn + j) Mod 4, i) & ".jpg")
            Else
                If (Placard((Turn + j) Mod 4, i) Mod 2) = 0 Then
                    ImgPla3(i) = LoadPicture(App.Path & "\img\B.jpg")
                Else
                    ImgPla3(i) = LoadPicture(App.Path & "\img\W.jpg")
                End If
            End If
        Else
            If PlacardB((Turn + j) Mod 4, i) = True Then
                ImgPla4(i) = LoadPicture(App.Path & "\img\" & Placard((Turn + j) Mod 4, i) & ".jpg")
            Else
                If (Placard((Turn + j) Mod 4, i) Mod 2) = 0 Then
                    ImgPla4(i) = LoadPicture(App.Path & "\img\B.jpg")
                Else
                    ImgPla4(i) = LoadPicture(App.Path & "\img\W.jpg")
                End If
            End If
        End If
    Next
Next

TimCardOp.Enabled = False
End Sub

Private Sub TimCardSort_Timer()
On Error GoTo err:
For a = 1 To GetTickCount / 1000
    Randomize ForPick
    ForPick = ((23 * Rnd) + 0)
Next
Exit Sub

err:
For a = 1 To 1000
    Randomize ForPick
    ForPick = ((23 * Rnd) + 0)
Next

End Sub
