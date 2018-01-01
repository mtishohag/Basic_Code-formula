VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   " EXIT"
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   " CONVERT "
      Height          =   975
      Left            =   2880
      TabIndex        =   4
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Text            =   " "
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Text            =   " "
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "AREA "
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   " CIRCLE"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim redious, circleArea As Double
Const PI = 3.1416
redious = Val(Text1)
circleArea = PI * radious * radious
Text2 = circleArea
End Sub

Private Sub Command2_Click()
End
End Sub
