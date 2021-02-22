VERSION 5.00
Begin VB.Form frmstart 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   360
      Top             =   2280
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   0
      Picture         =   "frmstart.frx":0000
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "frmstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
frmmain.Show
Me.Hide
Timer1.Enabled = False
End Sub
