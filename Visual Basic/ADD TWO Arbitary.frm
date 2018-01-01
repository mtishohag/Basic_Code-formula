VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008080&
   Caption         =   " Add two arbitary number"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   " Summation "
      Height          =   735
      Left            =   360
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   1440
      TabIndex        =   5
      Text            =   " "
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Text            =   "200"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Text            =   "100"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Result "
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " B"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " A"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b, result
a = Val(Text1)
b = Val(Text2)
result = a + b
Text3 = result
End Sub
