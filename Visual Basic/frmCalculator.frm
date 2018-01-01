VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   3615
   ClientLeft      =   7935
   ClientTop       =   3270
   ClientWidth     =   5130
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5130
   Begin VB.CommandButton cmdCalc 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   4020
      TabIndex        =   23
      Top             =   2940
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   3060
      TabIndex        =   22
      Top             =   2940
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   2100
      TabIndex        =   21
      Top             =   2940
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   1140
      TabIndex        =   20
      Top             =   2940
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   180
      TabIndex        =   19
      Top             =   2940
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   4020
      TabIndex        =   18
      Top             =   2400
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   3060
      TabIndex        =   17
      Top             =   2400
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   2100
      TabIndex        =   16
      Top             =   2400
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   1140
      TabIndex        =   15
      Top             =   2400
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   180
      TabIndex        =   14
      Top             =   2400
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   4020
      TabIndex        =   13
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   3060
      TabIndex        =   12
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   2100
      TabIndex        =   11
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1140
      TabIndex        =   10
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   180
      TabIndex        =   9
      Top             =   1860
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4020
      TabIndex        =   8
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3060
      TabIndex        =   7
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2100
      TabIndex        =   6
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1140
      TabIndex        =   5
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   180
      TabIndex        =   4
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1860
      TabIndex        =   2
      Top             =   720
      Width           =   1515
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   720
      Width           =   1515
   End
   Begin VB.Label lblDisplay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4755
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdblResult           As Double
Private mdblSavedNumber      As Double
Private mstrDot              As String
Private mstrOp               As String
Private mstrDisplay          As String
Private mblnDecEntered       As Boolean
Private mblnOpPending        As Boolean
Private mblnNewEquals        As Boolean
Private mblnEqualsPressed    As Boolean
Private mintCurrKeyIndex    As Integer

Private Sub Form_Load()

    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intIndex    As Integer
    
    Select Case KeyCode
        Case vbKeyBack:             intIndex = 0
        Case vbKeyDelete:           intIndex = 1
        Case vbKeyEscape:           intIndex = 2
        Case vbKey0, vbKeyNumpad0:  intIndex = 18
        Case vbKey1, vbKeyNumpad1:  intIndex = 13
        Case vbKey2, vbKeyNumpad2:  intIndex = 14
        Case vbKey3, vbKeyNumpad3:  intIndex = 15
        Case vbKey4, vbKeyNumpad4:  intIndex = 8
        Case vbKey5, vbKeyNumpad5:  intIndex = 9
        Case vbKey6, vbKeyNumpad6:  intIndex = 10
        Case vbKey7, vbKeyNumpad7:  intIndex = 3
        Case vbKey8, vbKeyNumpad8:  intIndex = 4
        Case vbKey9, vbKeyNumpad9:  intIndex = 5
        Case vbKeyDecimal:          intIndex = 20
        Case vbKeyAdd:              intIndex = 21
        Case vbKeySubtract:         intIndex = 16
        Case vbKeyMultiply:         intIndex = 11
        Case vbKeyDivide:           intIndex = 6
        Case Else:                  Exit Sub
    End Select
    
    cmdCalc(intIndex).SetFocus
    cmdCalc_Click intIndex
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim intIndex    As Integer
    
    Select Case Chr$(KeyAscii)
        Case "S", "s":  intIndex = 7
        Case "P", "p":  intIndex = 12
        Case "R", "r":  intIndex = 17
        Case "X", "x":  intIndex = 11
        Case "=":       intIndex = 22
        Case Else:      Exit Sub
    End Select
    
    cmdCalc(intIndex).SetFocus
    cmdCalc_Click intIndex

End Sub

Private Sub cmdCalc_Click(Index As Integer)

    Dim strPressedKey   As String
    
    mintCurrKeyIndex = Index
    
    If mstrDisplay = "ERROR" Then
        mstrDisplay = ""
    End If
    
    strPressedKey = cmdCalc(Index).Caption
    
    Select Case strPressedKey
        Case "0", "1", "2", "3", "4", _
             "5", "6", "7", "8", "9"
            If mblnOpPending Then
                mstrDisplay = ""
                mblnOpPending = False
            End If
            If mblnEqualsPressed Then
                mstrDisplay = ""
                mblnEqualsPressed = False
            End If
            mstrDisplay = mstrDisplay & strPressedKey
        Case "."
            If mblnOpPending Then
                mstrDisplay = ""
                mblnOpPending = False
            End If
            If mblnEqualsPressed Then
                mstrDisplay = ""
                mblnEqualsPressed = False
            End If
            If InStr(mstrDisplay, ".") > 0 Then
                Beep
            Else
                mstrDisplay = mstrDisplay & strPressedKey
            End If
        Case "+", "-", "X", "/"
            mdblResult = Val(mstrDisplay)
            mstrOp = strPressedKey
            mblnOpPending = True
            mblnDecEntered = False
            mblnNewEquals = True
        Case "%"
            mdblSavedNumber = (Val(mstrDisplay) / 100) * mdblResult
            mstrDisplay = Format$(mdblSavedNumber)
        Case "="
            If mblnNewEquals Then
                mdblSavedNumber = Val(mstrDisplay)
                mblnNewEquals = False
            End If
            Select Case mstrOp
                Case "+"
                    mdblResult = mdblResult + mdblSavedNumber
                Case "-"
                    mdblResult = mdblResult - mdblSavedNumber
                Case "X"
                    mdblResult = mdblResult * mdblSavedNumber
                Case "/"
                    If mdblSavedNumber = 0 Then
                        mstrDisplay = "ERROR"
                    Else
                        mdblResult = mdblResult / mdblSavedNumber
                    End If
                Case Else
                    mdblResult = Val(mstrDisplay)
            End Select
            If mstrDisplay <> "ERROR" Then
                mstrDisplay = Format$(mdblResult)
            End If
            mblnEqualsPressed = True
        Case "+/-"
            If mstrDisplay <> "" Then
                If Left$(mstrDisplay, 1) = "-" Then
                    mstrDisplay = Right$(mstrDisplay, 2)
                Else
                    mstrDisplay = "-" & mstrDisplay
                End If
            End If
        Case "Backspace"
            If Val(mstrDisplay) <> 0 Then
                mstrDisplay = Left$(mstrDisplay, Len(mstrDisplay) - 1)
                mdblResult = Val(mstrDisplay)
            End If
        Case "CE"
            mstrDisplay = ""
        Case "C"
            mstrDisplay = ""
            mdblResult = 0
            mdblSavedNumber = 0
        Case "1/x"
            If Val(mstrDisplay) = 0 Then
                mstrDisplay = "ERROR"
            Else
                mdblResult = Val(mstrDisplay)
                mdblResult = 1 / mdblResult
                mstrDisplay = Format$(mdblResult)
            End If
        Case "sqrt"
            If Val(mstrDisplay) < 0 Then
                mstrDisplay = "ERROR"
            Else
                mdblResult = Val(mstrDisplay)
                mdblResult = Sqr(mdblResult)
                mstrDisplay = Format$(mdblResult)
            End If
    End Select
        
    If mstrDisplay = "" Then
        lblDisplay = "0."
    Else
        mstrDot = IIf(InStr(mstrDisplay, ".") > 0, "", ".")
        lblDisplay = mstrDisplay & mstrDot
        If Left$(lblDisplay, 1) = "0" Then
            lblDisplay = Mid$(lblDisplay, 2)
        End If
    End If
    
    If lblDisplay = "." Then lblDisplay = "0."
    
End Sub

