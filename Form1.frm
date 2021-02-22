VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H00FFFFC0&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frameplay 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Play Board"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3495
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   5535
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   " &Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtply2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3480
      TabIndex        =   5
      Text            =   "O"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtply1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Text            =   "X"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H80000000&
      Caption         =   "&Start"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Now Turn of Player :"
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblname 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Computer :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Player 1  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "      Tic -Tac- Toe"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Menu fil 
      Caption         =   "&File"
      Begin VB.Menu filex 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu op 
      Caption         =   "&Option"
      Begin VB.Menu new 
         Caption         =   "&New game"
      End
      Begin VB.Menu opp 
         Caption         =   "Player Number"
         Begin VB.Menu opps 
            Caption         =   "&Singel Player"
         End
         Begin VB.Menu ppd 
            Caption         =   "&Double Player"
         End
      End
   End
   Begin VB.Menu el 
      Caption         =   "&Help"
      Begin VB.Menu elh 
         Caption         =   "About this game"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag, res As Boolean
Dim co As Integer


Private Sub Cmdnew_Click()
    Dim i As Integer
   ' Frameplay.Enabled = True
    For i = 0 To 8
        cmdplay(i).Caption = i
        cmdplay(i).BackColor = QBColor(15)
    Next i
    co = 0
    cmdstart.Enabled = True
    cmdnew.Enabled = False
    flag = False
End Sub

Private Sub cmdplay_Click(Index As Integer)
    On Error GoTo play_err
    If flag = True Then
        cmdplay(Index).Caption = txtply1.Text
        'cmdplay(Index).BackColor = QBColor(14)
        lblname.Caption = " 1 "
        flag = False
    Else
        cmdplay(Index).Caption = txtply2.Text
        'cmdplay(Index).BackColor = QBColor(13)
        lblname.Caption = " 2 "
        flag = True
    End If
    co = co + 1
    If cmdplay(0).Caption = cmdplay(1).Caption And cmdplay(1).Caption = cmdplay(2).Caption Then
     Call MsgBox("Player " & cmdplay(2).Caption & "  Win  ", vbOKOnly + vbInformation, "Result of the game ")
     res = False
     Frameplay.Enabled = False
     End If
     If cmdplay(3).Caption = cmdplay(4).Caption And cmdplay(4).Caption = cmdplay(5).Caption Then
     Call MsgBox("Player  " & cmdplay(4).Caption & "  Win ", vbOKOnly + vbInformation, "Result of the game ")
     res = False
     Frameplay.Enabled = False
     End If
     If cmdplay(6).Caption = cmdplay(7).Caption And cmdplay(7).Caption = cmdplay(8).Caption Then
     Call MsgBox("Player " & cmdplay(8).Caption & " Win ", vbOKOnly + vbInformation, "Result of the game ")
     res = False
     Frameplay.Enabled = False
     End If
     If cmdplay(0).Caption = cmdplay(3).Caption And cmdplay(3).Caption = cmdplay(6).Caption Then
     Call MsgBox("Player " & cmdplay(3).Caption & " Win ", vbOKOnly + vbInformation, "Result of the game ")
     res = False
     Frameplay.Enabled = False
     End If
     If cmdplay(1).Caption = cmdplay(4).Caption And cmdplay(4).Caption = cmdplay(7).Caption Then
     Call MsgBox("Player " & cmdplay(4).Caption & " Win ", vbOKOnly + vbInformation, "Result of the game ")
     res = False
     Frameplay.Enabled = False
     End If
     If cmdplay(2).Caption = cmdplay(5).Caption And cmdplay(5).Caption = cmdplay(8).Caption Then
     Call MsgBox("Player" & cmdplay(5).Caption & " Win ", vbOKOnly + vbInformation, "Result of the game ")
     res = False
     Frameplay.Enabled = False
     End If
     If cmdplay(0).Caption = cmdplay(4).Caption And cmdplay(4).Caption = cmdplay(8).Caption Then
     Call MsgBox("Player " & cmdplay(4).Caption & " Win ", vbOKOnly + vbInformation, "Result of the game ")
     res = False
     Frameplay.Enabled = False
     End If
     If cmdplay(2).Caption = cmdplay(4).Caption And cmdplay(4).Caption = cmdplay(6).Caption Then
     Call MsgBox("Player " & cmdplay(6).Caption & " Win ", vbOKOnly + vbInformation, "Result of the game ")
    res = False
    Frameplay.Enabled = False
    End If
    If res = False And co = 9 Then
    Call MsgBox("The game is Drow !!!!", vbInformation + vbOKOnly, "Information")
    Frameplay.Enabled = False
    End If
    Exit Sub
play_err:
Call MsgBox(Err.Description, vbOKOnly + 32, "Error massage")
End Sub

Private Sub cmdstart_Click()
Dim i As Integer
   Frameplay.Enabled = True
   cmdnew.Enabled = True
   cmdstart.Enabled = False
   lblname.Caption = " 1 "

    Frameplay.Enabled = True
    For i = 0 To 8
        cmdplay(i).Caption = i
    cmdplay(i).BackColor = QBColor(15)
    Next i
    co = 0
    flag = False
  Refresh
End Sub

Private Sub Command1_Click()
    Unload Me
    End
End Sub

Private Sub elh_Click()
Beep
frmma.Show
Me.Hide
End Sub

Private Sub filex_Click()
End
End Sub

Private Sub Form_Load()
Dim name As String
flag = True
frmmain.Show
name = InputBox("Please Enter your Name : ", "Name of the Player1")
Label2.Caption = UCase(name)
cmdnew.Enabled = False
Dim name1 As String
flag = True
Beep
name1 = InputBox("Please Enter your Name : ", "Name of the Player2")
Label3.Caption = UCase(name1)

End Sub


Private Sub new_Click()
Dim i As Integer
    'Frameplay.Enabled = True
    For i = 0 To 8
        cmdplay(i).Caption = i
    Next i
    co = 0
    flag = False
Refresh
End Sub

Private Sub opps_Click()
Dim name As String
flag = True
UCase(name) = InputBox("Please Enter your Name : ", "Name of the Player1")
Label2.Caption = UCase(name)
Label3.Caption = UCase("Computer")
End Sub

Private Sub ppd_Click()
Dim name1 As String
flag = True
UCase(name1) = InputBox("Please Enter your Name : ", "Name of the Player2")
Label3.Caption = UCase(name1)
End Sub

