VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " Even no "
   ClientHeight    =   3015
   ClientLeft      =   7950
   ClientTop       =   4305
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   " Print even no  "
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a
For a = 2 To 50 Step 2
Print a
Next a
End Sub
