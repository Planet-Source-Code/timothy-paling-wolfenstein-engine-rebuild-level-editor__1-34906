VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load Balancing"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label6 
      Caption         =   "Use this slider to alter the sensitivity of the viewpoint"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "Sensitive"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Insensitive"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   4560
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Insensitive"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Sensitive"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Use this slider to alter the sensitivity of the puck."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form7"
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

