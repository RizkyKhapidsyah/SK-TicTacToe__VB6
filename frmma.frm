VERSION 5.00
Begin VB.Form frmma 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "About the  Tic Tac Toe (1.0) game"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   FillColor       =   &H80000008&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label6 
      BackColor       =   &H00404040&
      Caption         =   "Mail  To ....."
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "napollab@yahoo.com"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2400
      Picture         =   "frmma.frx":030A
      ToolTipText     =   "Tic Tac Toe (1.0)"
      Top             =   -600
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "    Departments of Computer Science  and                           Engineering   (C.S.E)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5775
   End
End
Attribute VB_Name = "frmma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Image2_Click()
Me.Hide
frmmain.Show
End Sub

Private Sub Label5_Click()
con = Shell("start mailto:" + "nalam@vasdigital.com?subject=Shuffle%201.0%20Comments", vbMinimizedFocus)
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = QBColor(15)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = QBColor(9)
End Sub

