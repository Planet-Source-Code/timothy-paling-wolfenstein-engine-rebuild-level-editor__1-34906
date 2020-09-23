VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KeyMapper"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   495
      Left            =   840
      TabIndex        =   30
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   29
      Text            =   "79"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   28
      Text            =   "82"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Text            =   "211"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   26
      Text            =   "44"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   25
      Text            =   "30"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Text            =   "33"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Text            =   "31"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Text            =   "32"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   21
      Text            =   "18"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   20
      Text            =   "207"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Text            =   "199"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   18
      Text            =   "203"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Text            =   "205"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Text            =   "208"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Text            =   "200"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Rotate Puck:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Place Selected:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Delete:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "View Look Down:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "View Look Up:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "View Rotate Right:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "View Rotate Left:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "View B'ward:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "View forwards:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Puck Z-:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Puck Z+:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Puck Y-:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Puck Y+:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Puck X-:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Puck X+:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form4"
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
End Sub
