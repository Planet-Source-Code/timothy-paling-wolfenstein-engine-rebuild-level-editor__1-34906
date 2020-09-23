VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directory Selection"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   1695
      Left            =   120
      Top             =   720
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
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
Private Sub Command1_Click()
    Me.Hide
    Form1.File1.Path = Form2.Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    
End Sub
