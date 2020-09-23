VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Environment Settings"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Apply Ambient"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5520
      TabIndex        =   13
      Text            =   "&FFFFFFFF"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply Fog"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Use Range Fog"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "&HFFFFFFFF"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtFogDensity 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "0"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtFogEnd 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "2"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtFogStart 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "0.1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Ambient Colour [Hex]:"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   3375
      Left            =   3720
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Mode:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Colour [Hex]:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Fog Density:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Fog End:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fog Start:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   3375
      Left            =   120
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form5"
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
    Dim fMode As Integer
    Select Case Combo1.Text
        Case "D3D_FOGLINEAR"
            SetupVertexFog Text1.Text, D3DFOG_LINEAR, Check1.Value, txtFogStart.Text, txtFogEnd.Text, txtFogDensity.Text
        Case "D3D_FOGEXP"
            SetupVertexFog Text1.Text, D3DFOG_EXP, Check1.Value, txtFogStart.Text, txtFogEnd.Text, txtFogDensity.Text
        Case "D3D_FOGEXP2"
            SetupVertexFog Text1.Text, D3DFOG_EXP2, Check1.Value, txtFogStart.Text, txtFogEnd.Text, txtFogDensity.Text
    End Select
    Me.Hide
    
End Sub

Private Sub Command2_Click()
    g_D3DDevice.SetRenderState D3DRS_AMBIENT, Text2.Text
    Me.Hide
End Sub

Private Sub Form_Load()
    Combo1.AddItem "D3D_FOGLINEAR"
    Combo1.AddItem "D3D_FOGEXP"
    Combo1.AddItem "D3D_FOGEXP2"
    Combo1.Text = "D3D_FOGLINEAR"
End Sub
