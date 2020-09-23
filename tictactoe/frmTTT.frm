VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm2Player 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Two Player Tic Tac Toe Fpr Intra/Inter net"
   ClientHeight    =   3105
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Winsock1 
      Height          =   480
      Left            =   4320
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   12
      Top             =   3720
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4320
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4320
      Top             =   3000
   End
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   3600
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsServer 
      Index           =   0
      Left            =   3240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Designed By Shyam Bhand and Dadu Ritesh Nath Singh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP Address"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   7
      X1              =   240
      X2              =   2400
      Y1              =   2280
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   6
      X1              =   360
      X2              =   2400
      Y1              =   240
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   5
      X1              =   2040
      X2              =   2040
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   4
      X1              =   1320
      X2              =   1320
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   3
      X1              =   480
      X2              =   480
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   2
      X1              =   240
      X2              =   2520
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   1
      X1              =   240
      X2              =   2520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   0
      X1              =   240
      X2              =   2520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label turn 
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   8
      Left            =   1920
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   7
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   5
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   240
      X2              =   2520
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   240
      X2              =   2520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   1800
      X2              =   1800
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   960
      X2              =   960
      Y1              =   120
      Y2              =   2400
   End
End
Attribute VB_Name = "frm2Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------
'---------------------------
'This special game is designed for the Employees
'whose Employer doesn't give them enough work to get on.
'These are a class of Employees who are having Internet as well as Intranet facilities.
'---------------------------
'---------------------------
Dim iwon As Boolean
Dim cwon As Boolean
Dim tie As Boolean
Dim Flag1 As Boolean
Dim Flag As Boolean
Dim Flag2 As Boolean
Private Sub cmdConnect_Click()
    Unload Me
    frmConnection.Show
End Sub

Private Sub Command1_Click()
    Dim a As Integer
    Call New_Game
    a = 90
    wsClient.SendData a
End Sub

Private Sub Command2_Click()
    Call Quiter
End Sub

Private Sub Command3_Click()
    wsClient.Connect Text1.Text, 5555
End Sub

Private Sub Form_Load()
    wsServer(0).LocalPort = 5555
    wsServer(0).Listen
    turn.Caption = "X"
    Line5(0).Visible = False
    For c = 1 To 7
        h = h + 1
        Line5(h).Visible = False
    Next c
End Sub

Private Sub s_Click(y As Integer)
    If turn.Caption = "X" And s(y).Caption = "" And cwon = False And iwon = False And tie = False Then
        s(y).Caption = "X"
        turn.Caption = "O"
        wsClient.SendData y
    End If
    If turn.Caption = "O" And s(y).Caption = "" And cwon = False And iwon = False And tie = False Then
        s(y).Caption = "O"
        turn.Caption = "X"
        wsClient.SendData y
    End If

    If turn.Caption = "X" And s(y).Caption = "" And cwon = False And iwon = False And tie = False Then
        s(y).Caption = "X"
        turn.Caption = "O"
        wsClient.SendData y
    End If
    If turn.Caption = "O" And s(y).Caption = "" And cwon = False And iwon = False And tie = False Then
        s(y).Caption = "O"
        turn.Caption = "X"
        wsClient.SendData y
    End If
    
End Sub


Private Sub Timer1_Timer()

If s(0).Caption = "X" And s(1).Caption = "X" And s(2).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(0).Visible = True
End If

If s(3).Caption = "X" And s(4).Caption = "X" And s(5).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(1).Visible = True
End If

If s(6).Caption = "X" And s(7).Caption = "X" And s(8).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(2).Visible = True

End If

If s(0).Caption = "X" And s(3).Caption = "X" And s(6).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(3).Visible = True


End If

If s(1).Caption = "X" And s(4).Caption = "X" And s(7).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(4).Visible = True

End If

If s(2).Caption = "X" And s(5).Caption = "X" And s(8).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(5).Visible = True

End If

If s(0).Caption = "X" And s(4).Caption = "X" And s(8).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(6).Visible = True

End If

If s(2).Caption = "X" And s(4).Caption = "X" And s(6).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(7).Visible = True
End If

If s(0).Caption = "O" And s(1).Caption = "O" And s(2).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(0).Visible = True

End If

If s(3).Caption = "O" And s(4).Caption = "O" And s(5).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(1).Visible = True

End If

If s(6).Caption = "O" And s(7).Caption = "O" And s(8).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(2).Visible = True

End If

If s(0).Caption = "O" And s(3).Caption = "O" And s(6).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(3).Visible = True

End If

If s(1).Caption = "O" And s(4).Caption = "O" And s(7).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(4).Visible = True
End If

If s(2).Caption = "O" And s(5).Caption = "O" And s(8).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(5).Visible = True


End If

If s(0).Caption = "O" And s(4).Caption = "O" And s(8).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(6).Visible = True

End If

If s(2).Caption = "O" And s(4).Caption = "O" And s(6).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(7).Visible = True
End If

If s(0) <> "" And s(1) <> "" And s(2) <> "" And s(3) <> "" And s(4) <> "" And s(5) <> "" And s(6) <> "" And s(7) <> "" And s(8) <> "" And cwon = False And iwon = False Then
tie = True
iwon = False
cwon = False
End If
End Sub

Private Sub Timer2_Timer()
If iwon = True Then
Timer1.Interval = 0
Timer2.Interval = 0
x = MsgBox("X's win.", vbOKOnly, "Tic Tac Toe")
End If
If cwon = True Then
Timer1.Interval = 0
Timer2.Interval = 0
o = MsgBox("O's win.", vbOKOnly, "Tic Tac Toe")

End If
If tie = True Then
Timer1.Interval = 0
Timer2.Interval = 0
ci = MsgBox("There is no winner.  It's a tie.", vbOKOnly, "Tic Tac Toe")
End If
End Sub

Private Sub Quiter()
Unload Me
End Sub

Private Sub New_Game()
s(0).Caption = ""
For Index = 1 To 8
    num = num + 1
    s(num).Caption = ""
Next Index
Line5(0).Visible = False
For i = 1 To 7
    x = x + 1
    Line5(x).Visible = False
Next i
Timer1.Interval = 1
Timer2.Interval = 1
cwon = False
iwon = False
tie = False
turn.Caption = "X"
End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
    Dim strDataRecived As Integer
    wsClient.GetData strDataRecived
    DoEvents
    s_Click (strDataRecived)
End Sub

Private Sub wsServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Load wsServer(Index + 1)
wsServer(Index + 1).Accept requestID
End Sub

Private Sub wsServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strRecivedData As Integer
    Dim SocketCheck As Integer
    wsServer(Index).GetData strRecivedData
    If strRecivedData = 90 Then
        Call New_Game
    Else
        Flag1 = True
        s_Click (strRecivedData)
    End If
End Sub
