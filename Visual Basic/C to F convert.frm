VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   Caption         =   " C to F convert"
   ClientHeight    =   5775
   ClientLeft      =   7905
   ClientTop       =   4635
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   10500
   Begin VB.CommandButton Command2 
      Caption         =   " Exit"
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   " C T O F CONVERT"
      Height          =   1215
      Left            =   3360
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   3600
      TabIndex        =   3
      Text            =   " "
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   " FARENHEIGHT"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   " CELCIUS "
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim celcius, farenheight As Double
celcius = Val(Text1)
farenheight = (1.8 * celcius) + 32
Text2 = farenheight
End Sub

Private Sub Command2_Click()
End
End Sub
