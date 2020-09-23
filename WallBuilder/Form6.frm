VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Environmetal Effects"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4140
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox V4Y 
      Height          =   285
      Left            =   2160
      TabIndex        =   23
      Text            =   "0"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox V3Y 
      Height          =   285
      Left            =   2160
      TabIndex        =   21
      Text            =   "0"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox V2Y 
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Text            =   "0"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox V1Y 
      Height          =   285
      Left            =   2160
      TabIndex        =   17
      Text            =   "0"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox v4blue 
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Text            =   "255"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox v4green 
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Text            =   "255"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox v3blue 
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Text            =   "255"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox v3green 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Text            =   "255"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox v2blue 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Text            =   "255"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox v2green 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "255"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox v1blue 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "255"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox v1green 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "255"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox v4red 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "255"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox v3red 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "255"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox v2red 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "255"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox v1red 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "255"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 4 Y:"
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 3 Y:"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 2 Y:"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 1 Y:"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 4:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 3:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 2:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 1:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form6"
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
    Text2.Text = Text1.Text
    Text3.Text = Text1.Text
    Text4.Text = Text1.Text
End Sub
