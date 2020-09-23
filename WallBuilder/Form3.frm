VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10980
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   240
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   6000
      ScaleWidth      =   10500
      TabIndex        =   0
      Top             =   240
      Width           =   10500
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2880
         Top             =   2760
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Level Editor - (c)2001 Timothy Paling. This is a pre-release version. Please do not redistribute."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   6360
      Width           =   6735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()
    Me.Hide
    Timer1.Enabled = False
    Form1.Show
End Sub
